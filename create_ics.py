import csv
import glob
import os
import re
import sys
from datetime import datetime

import pandas as pd
import pytz
from icalendar import Calendar, Event, vRecur


# Load address map from CSV
def load_addresses(csv_path):
    address_map = {}
    with open(csv_path, "r") as file:
        reader = csv.reader(file)
        next(reader)  # skip header if present
        for row in reader:
            building_name, code, address = row
            code = code.strip().upper()
            address_map[code] = {
                "name": building_name.strip(),
                "address": address.strip(),
            }
    return address_map


ADDRESS_MAP = load_addresses("Address.csv")


# Function to parse location and append full address
def parse_address(location, address_map):
    if not location or not str(location).strip():
        return "Unknown Address"
    loc = str(location).strip()

    if loc.lower().startswith("online"):
        return "Online - Virtual Class\nCanada"

    # Try new format: "... (HENN) | ... | Room: 200"
    m_code = re.search(r"\(([A-Z0-9]+)\)", loc)
    m_room = re.search(r"Room:\s*([A-Za-z0-9\-]+)", loc)

    code = m_code.group(1).upper() if m_code else None
    room = m_room.group(1) if m_room else None

    # Fallback to old format: "HENN - Room 200"
    if not code:
        parts = loc.split("-")
        if parts:
            code = parts[0].strip().upper()
        if len(parts) >= 2 and not room:
            room = parts[-1].strip().replace("Room", "").strip()

    if code in address_map:
        addr = address_map[code]["address"]
        if room:
            return f"{room}-{addr}\nVancouver BC\nCanada"
        return f"{addr}\nVancouver BC \nCanada"

    return "Unknown Address"


# Function to get the full building name
def get_building_full_name(location, address_map):
    if not location or not str(location).strip():
        return str(location)

    loc = str(location).strip()
    if loc.lower().startswith("online"):
        return "💻 Online Class"

    # New format: "... | Hennings Building (HENN) | Floor: 1 | Room: 200"
    m_code = re.search(r"\(([A-Z0-9]+)\)", loc)
    code = m_code.group(1).upper() if m_code else None
    pieces = [p.strip() for p in loc.split("|")]
    human_name = None
    if len(pieces) >= 2:
        # remove trailing "(CODE)" from the building name
        human_name = re.sub(r"\s*\([A-Z0-9]+\)\s*$", "", pieces[1]).strip()

    m_room = re.search(r"Room:\s*([A-Za-z0-9\-]+)", loc)
    room = m_room.group(1) if m_room else None

    if code in address_map:
        name = human_name or address_map[code]["name"]
        if room:
            return f"📍{name} ({code}) - Room {room}"
        return f"📍{name} ({code})"

    # Fallbacks
    if code and room:
        return f"📍{code} - Room {room}"
    return loc


# Function to parse time
def parse_time(time_str):
    time_str = time_str.lower().replace(".", "").replace("|", "").strip()
    return datetime.strptime(time_str, "%I:%M %p").time()


# Function to parse meeting patterns
def parse_meeting_pattern(pattern):
    parts = pattern.strip().split(" | ")

    if len(parts) < 3:
        raise ValueError(f"Invalid pattern format: {pattern}")

    # Always present
    dates = parts[0].strip()
    days = parts[1].strip()
    times = parts[2].strip()

    # Flexible location handling
    if len(parts) >= 4:
        location_parts = parts[3:]
        location = " | ".join(location_parts).strip()
    else:
        location = "Online"

    # Check for alternating weeks
    is_alternating = False
    if "alternate" in days.lower() or "alternating" in days.lower():
        is_alternating = True
        # Remove the alternating weeks text
        days = re.sub(
            r"\s*\(alternate\s+weeks\)|\(alternating\s+weeks\)",
            "",
            days,
            flags=re.IGNORECASE,
        ).strip()

    # Time parsing
    start_time, end_time = map(parse_time, times.split(" - "))
    start_date, end_date = map(
        lambda x: datetime.strptime(x.strip(), "%Y-%m-%d").date(), dates.split(" - ")
    )

    return (
        start_date,
        end_date,
        days.split(),
        start_time,
        end_time,
        location,
        is_alternating,
    )


# Function to create an event
def create_event(
    name, start_datetime, end_datetime, location, address="", description=""
):
    event = Event()
    event.add("summary", name)
    event.add("dtstart", start_datetime)
    event.add("dtend", end_datetime)
    event.add("location", address)
    event.add("description", description)
    return event


def process_schedule(xlsx_path, output_path):
    """Process Excel schedule and generate ICS calendar file."""
    print(f"Processing file: {xlsx_path}")

    # Process the file and generate the .ics
    df_initial = pd.read_excel(xlsx_path, header=None)
    row_with_course_listing = None
    for i, row in df_initial.iterrows():
        if "Course Listing" in row.values:
            row_with_course_listing = i
            break

    if row_with_course_listing is None:
        print("Error: Could not find 'Course Listing' in the Excel file.")
        return False

    df = pd.read_excel(xlsx_path, skiprows=row_with_course_listing)

    cal = Calendar()
    days_map = {"Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3, "Fri": 4, "Sat": 5, "Sun": 6}
    events_created = 0

    for _, row in df.iterrows():
        if (
            pd.isna(row.get("Meeting Patterns"))
            or not str(row["Meeting Patterns"]).strip()
        ):
            continue
        # Extract course code (before first " - ") and combine with format
        course_listing = str(row["Course Listing"])
        course_code = course_listing.split(" - ")[0].strip()
        name = f"{course_code} - {row['Instructional Format']}"
        meeting_patterns = row["Meeting Patterns"]
        instructor = row.get("Instructor", "")
        details = row.get("Section", "")
        patterns = re.split(r"\n(?=\d{4})", meeting_patterns)

        for pattern in patterns:
            (
                start_date,
                end_date,
                days,
                start_time,
                end_time,
                location,
                is_alternating,
            ) = parse_meeting_pattern(pattern)
            weekday_map = {
                "Mon": "MO",
                "Tue": "TU",
                "Wed": "WE",
                "Thu": "TH",
                "Fri": "FR",
                "Sat": "SA",
                "Sun": "SU",
            }
            byday = [weekday_map[day] for day in days if day in weekday_map]

            # Set timezone to Vancouver
            vancouver_tz = pytz.timezone("America/Vancouver")
            start_datetime = vancouver_tz.localize(
                datetime.combine(start_date, start_time)
            )
            end_datetime = vancouver_tz.localize(datetime.combine(start_date, end_time))
            full_address = parse_address(location, ADDRESS_MAP)
            building_full_description = get_building_full_name(location, ADDRESS_MAP)
            description = (
                f"Instructor: {instructor}\n\n{building_full_description}\n\n{details}"
            )

            event = create_event(
                name, start_datetime, end_datetime, location, full_address, description
            )

            # Set up recurrence rule with interval for alternating weeks
            rrule_params = {
                "freq": "weekly",
                "byday": byday,
                "until": vancouver_tz.localize(datetime.combine(end_date, end_time)),
            }
            if is_alternating:
                rrule_params["interval"] = 2

            event.add("rrule", vRecur(rrule_params))
            cal.add_component(event)
            events_created += 1

    # Write the ICS file
    with open(output_path, "wb") as f:
        f.write(cal.to_ical())

    print(f"✓ Created {events_created} calendar events")
    print(f"✓ Calendar saved to: {output_path}")
    return True


def main():
    """Main function to run the schedule converter."""
    print("Schedule to Calendar Converter")
    print("=" * 50)

    # Find xlsx files in current directory
    xlsx_files = glob.glob("*.xlsx")

    if not xlsx_files:
        print("Error: No .xlsx files found in the current directory")
        sys.exit(1)

    if len(xlsx_files) == 1:
        input_file = xlsx_files[0]
        print(f"Found schedule file: {input_file}")
    else:
        print(f"Found {len(xlsx_files)} .xlsx files:")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  {i}. {f}")

        while True:
            try:
                choice = input("\nEnter the number of the file to process: ")
                idx = int(choice) - 1
                if 0 <= idx < len(xlsx_files):
                    input_file = xlsx_files[idx]
                    break
                else:
                    print(f"Please enter a number between 1 and {len(xlsx_files)}")
            except (ValueError, KeyboardInterrupt):
                print("\nOperation cancelled")
                sys.exit(0)

    # Generate output filename
    base_name = os.path.splitext(input_file)[0]
    output_file = f"{base_name}.ics"

    print()
    # Process the schedule
    success = process_schedule(input_file, output_file)

    if success:
        print("\n" + "=" * 50)
        print("Conversion complete!")
    else:
        print("\nConversion failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
