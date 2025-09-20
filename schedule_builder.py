import csv
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple
import re

class ScheduleBuilder:
    def __init__(self):
        self.time_slots = []
        self.groups = {}
        self.schedule_data = []

    def parse_hierarchy(self, group_name: str) -> Tuple[List[str], str]:
        """
        Parse hierarchical group names like "Players - Planners - AOs"
        Returns tuple of (hierarchy_levels, leaf_name)
        """
        if ' - ' in group_name:
            parts = [part.strip() for part in group_name.split(' - ')]
            return parts[:-1], parts[-1]
        else:
            return [], group_name

    def generate_time_slots(self, start_time: str, end_time: str, interval_minutes: int = 30) -> List[str]:
        """
        Generate time slots between start_time and end_time
        Format: "HHMM" (24-hour format)
        """
        start_hour = int(start_time[:2])
        start_min = int(start_time[2:])

        # Handle 2400 as midnight (0000 of next day)
        if end_time == "2400":
            end_hour = 23
            end_min = 59
        else:
            end_hour = int(end_time[:2])
            end_min = int(end_time[2:])

        start_dt = datetime(2000, 1, 1, start_hour, start_min)
        end_dt = datetime(2000, 1, 1, end_hour, end_min)

        # Handle overnight schedules or 2400 case
        if end_dt <= start_dt or end_time == "2400":
            end_dt += timedelta(days=1)

        slots = []
        current = start_dt
        while current <= end_dt:
            time_str = current.strftime("%H%M")
            # Convert midnight to 2400 format if needed
            if time_str == "0000" and current > start_dt:
                time_str = "2400"
            slots.append(time_str)
            current += timedelta(minutes=interval_minutes)

        return slots

    def add_group(self, group_name: str, activities: Dict[str, str] = None, locations: List[str] = None):
        """
        Add a group with optional activities and locations
        group_name: hierarchical name like "Players - Planners - AOs"
        activities: dict mapping time slots to activity descriptions
        locations: list of locations for this group
        """
        hierarchy, leaf_name = self.parse_hierarchy(group_name)

        if group_name not in self.groups:
            self.groups[group_name] = {
                'hierarchy': hierarchy,
                'leaf_name': leaf_name,
                'activities': activities or {},
                'locations': locations or []
            }

    def set_time_period(self, start_time: str, end_time: str, interval_minutes: int = 30):
        """
        Set the time period for the schedule
        """
        self.time_slots = self.generate_time_slots(start_time, end_time, interval_minutes)

    def add_activity(self, group_name: str, time_slot: str, activity: str, location: str = ""):
        """
        Add an activity for a specific group at a specific time
        """
        if group_name not in self.groups:
            self.add_group(group_name)

        self.groups[group_name]['activities'][time_slot] = activity
        if location and location not in self.groups[group_name]['locations']:
            self.groups[group_name]['locations'].append(location)

    def build_schedule_structure(self) -> List[List[str]]:
        """
        Build the hierarchical schedule structure similar to the existing CSV format
        """
        if not self.time_slots:
            raise ValueError("Time period must be set before building schedule")

        # Create header row with time slots
        header = ['', '', ''] + self.time_slots
        schedule = [header, ['', '', ''] + [''] * len(self.time_slots)]  # Empty row

        # Group by hierarchy levels
        hierarchy_groups = {}
        for group_name, group_data in self.groups.items():
            hierarchy = group_data['hierarchy']
            leaf = group_data['leaf_name']

            # Create hierarchy key
            if not hierarchy:
                # Top-level group
                key = (group_name,)
            else:
                key = tuple(hierarchy + [leaf])

            hierarchy_groups[key] = group_data

        # Sort groups by hierarchy
        sorted_groups = sorted(hierarchy_groups.items())

        current_top_level = None
        current_second_level = None

        for group_key, group_data in sorted_groups:
            hierarchy = group_data['hierarchy']
            leaf = group_data['leaf_name']
            activities = group_data['activities']
            locations = group_data['locations']

            # Determine the row structure based on hierarchy depth
            if len(group_key) == 1:
                # Top-level group
                current_top_level = group_key[0]
                current_second_level = None
                row = [current_top_level, '', '']
            elif len(group_key) == 2:
                # Second-level group
                top_level = group_key[0] if group_key[0] != current_top_level else ''
                current_top_level = group_key[0]
                current_second_level = group_key[1]
                row = [top_level, current_second_level, '']
            elif len(group_key) == 3:
                # Third-level group
                top_level = group_key[0] if group_key[0] != current_top_level else ''
                second_level = group_key[1] if group_key[1] != current_second_level else ''
                current_top_level = group_key[0]
                current_second_level = group_key[1]
                row = [top_level, second_level, group_key[2]]
            else:
                # Handle deeper hierarchies by using the last level as the third column
                row = ['', '', leaf]

            # Add location info if available
            if locations:
                if row[2]:
                    row[2] += f" ({', '.join(locations)})"
                else:
                    row[2] = f"({', '.join(locations)})"

            # Fill in activities for each time slot
            for time_slot in self.time_slots:
                activity = activities.get(time_slot, '')
                row.append(activity)

            schedule.append(row)

        return schedule

    def export_to_csv(self, filename: str):
        """
        Export the schedule to a CSV file
        """
        schedule = self.build_schedule_structure()

        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            for row in schedule:
                writer.writerow(row)

    def export_to_dict(self) -> Dict:
        """
        Export the schedule as a dictionary for programmatic use
        """
        return {
            'time_slots': self.time_slots,
            'groups': self.groups,
            'schedule': self.build_schedule_structure()
        }

def create_example_schedule() -> ScheduleBuilder:
    """
    Create an example schedule to demonstrate the functionality
    """
    builder = ScheduleBuilder()

    # Set time period from 0530 to 2400 with 30-minute intervals
    builder.set_time_period("0530", "2400", 30)

    # Add hierarchical groups with activities
    builder.add_group("Players - Echelon 2 & 3 - CPF")
    builder.add_activity("Players - Echelon 2 & 3 - CPF", "0630", "SAP CUB")
    builder.add_activity("Players - Echelon 2 & 3 - CPF", "0700", "TS Cmdrs Update Brief (CUB)")
    builder.add_activity("Players - Echelon 2 & 3 - CPF", "0800", "Supervise Mission Creation")

    builder.add_group("Players - Echelon 2 & 3 - Commanders")
    builder.add_activity("Players - Echelon 2 & 3 - Commanders", "1130", "Submit Msns")
    builder.add_activity("Players - Echelon 2 & 3 - Commanders", "1200", "Lunch")

    builder.add_group("Players - Planners - Leads")
    builder.add_activity("Players - Planners - Leads", "0700", "CUB")
    builder.add_activity("Players - Planners - Leads", "0800", "Direct Planning")

    return builder

if __name__ == "__main__":
    # Example usage
    schedule = create_example_schedule()
    schedule.export_to_csv("example_schedule.csv")
    print("Example schedule created as 'example_schedule.csv'")