# UBC Class Schedule to Calendar

> This fork is a minimal version of [pouyanfz/Schedule2Calendar](https://github.com/pouyanfz/Schedule2Calendar) that foregoes the Flask server in favour of command-line interactivity. It also features proper logic for handling alternate week course entries.

A command-line tool that allows UBC students to seamlessly convert their class schedules into calendar events with full details, including the class location and the option to set notifications to remind them when to leave for class. This tool is especially helpful for those who find it difficult to navigate the campus, and it automates the process of tracking your classes while also including useful location data for easy navigation.

## Motivation

As a student at UBC, I found it frustrating when I first arrived, not knowing how to navigate the campus and how far my next class might be. I often had trouble finding classrooms, and this led to unnecessary stress. With this tool, I wanted to create a solution that provides all the necessary information—class times, locations, and reminders—directly in your calendar.

Now, you can simply upload your class schedule, and the application will generate a `.ics` calendar file that includes all the relevant details, so you don't have to worry about finding your classes again. It even lets you set notifications to remind you when to leave based on the time it takes to walk between buildings!

## Features

- **Class Schedule to Calendar**: Convert your class schedule into calendar events with full details.
- **Location and Address**: Automatically adds the address of your classrooms and provides a full location for easy navigation.
- **Time to Leave Notifications**: Set reminders to alert you when it's time to leave for your next class based on your location.
- **ICS File Export**: Export the schedule directly into an ICS file that can be imported into your preferred calendar application (Google Calendar, Apple Calendar, etc.).

## How It Works

1. Download your class schedule from Workday: Navigate to **Academics -> Registration and Courses**, click the ⚙️ in the current class tab, and select **Download to Excel**.
2. Place the downloaded Excel file in the project directory.
3. Run the script - it will automatically detect your schedule file.
4. The application reads the file, extracts the necessary information (course names, times, locations), and generates a `.ics` calendar file.
5. Import the generated `.ics` file into your calendar application (Google Calendar, Apple Calendar, etc.).

### Example of Class Event

- **Course Name**: CPSC_V 213 - Introduction to Computer Systems
- **Location**: Room 101-6245 Agronomy Road Vancouver BC, Canada
- **Instructor**: Rubeus Hagrid
- **Time**: Monday, Wednesday, Friday from 14:00 PM to 15:30 PM

### Notifications
If you want to get **time to leave** notifications on iOS you can navigate to **Settings -> Apps -> Calendar -> Default Alert Time** and set the **Time to Leave** reminder so you don't miss your class. The application will notify you with enough time to navigate across campus and get to your classroom on time.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/munr0/schedule2calendar.git
   cd schedule2calendar
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv

   # On Windows
   .\venv\Scripts\Activate.ps1

   # On macOS/Linux
   source venv/bin/activate
   ```

3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Place your Excel schedule file (e.g., `View_My_Courses.xlsx`) in the project directory.

2. Run the script:
   ```bash
   python create_ics.py
   ```

3. The script will automatically find your schedule file and generate a `.ics` file with the same name.

4. Import the generated `.ics` file into your calendar application.
