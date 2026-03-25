# -*- coding: utf-8 -*-
"""
HR Scheduler - Human Resource Scheduling Tool
Author: Assistant
"""

import sys
import os
from datetime import datetime, timedelta
from collections import defaultdict
from typing import Dict, List, Set, Tuple, Optional
from dataclasses import dataclass, field
import argparse
import json

# Fix Windows encoding
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    print("Warning: pandas not installed, using built-in CSV support")


@dataclass
class Person:
    """Person information"""
    name: str
    skills: Dict[str, int] = field(default_factory=dict)
    available_from: datetime = None
    available_until: datetime = None
    
    def has_skill(self, skill: str, min_level: int = 1) -> bool:
        return self.skills.get(skill, 0) >= min_level
    
    def skill_score(self, required_skills: Dict[str, int]) -> float:
        if not required_skills:
            return 0.5
        total = 0.0
        for skill, min_level in required_skills.items():
            level = self.skills.get(skill, 0)
            if level >= min_level:
                total += level / 5.0
        return total / len(required_skills)


@dataclass
class Assignment:
    """Assignment record"""
    person_name: str
    month: str
    role: str = "Test"
    skills_used: List[str] = field(default_factory=list)
    is_original: bool = False


@dataclass
class BaselineRequirement:
    """Baseline requirement"""
    month: str
    required_count: int
    required_skills: Dict[str, int] = field(default_factory=dict)


class HRScheduler:
    """HR Scheduler"""
    
    def __init__(self):
        self.resource_pool: Dict[str, Person] = {}
        self.baseline: List[BaselineRequirement] = []
        self.original_assignments: List[Assignment] = []
        self.test_features: Dict[str, Dict[str, int]] = {}
    
    def parse_skills(self, skill_str: str) -> Dict[str, int]:
        """Parse skill string like 'Python:3,Java:4'"""
        skills = {}
        if not skill_str or skill_str == 'nan':
            return skills
        for item in str(skill_str).split(','):
            item = item.strip()
            if ':' in item:
                parts = item.split(':')
                if len(parts) == 2:
                    skill, level = parts
                    try:
                        skills[skill.strip()] = int(level.strip())
                    except ValueError:
                        pass
        return skills
    
    def load_resource_pool(self, file_path: str) -> None:
        """Load resource pool from Excel or CSV"""
        if HAS_PANDAS:
            df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
        else:
            import csv
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                df = list(reader)
        
        for row in (df.iterrows() if HAS_PANDAS else enumerate(df)):
            row_data = row[1] if HAS_PANDAS else row
            name = str(row_data.get('name', row_data.get('Name', row_data.get('name', ''))))
            if not name or name == 'nan':
                continue
            
            skills = {}
            for key, val in (row_data.items() if HAS_PANDAS else row_data.items()):
                key_lower = key.lower()
                if key_lower not in ['name', 'name', 'available_from', 'available_until', 'available from', 'available until']:
                    try:
                        if HAS_PANDAS:
                            if pd.notna(val) and float(val) > 0:
                                skills[key] = int(float(val))
                        else:
                            if val and float(val) > 0:
                                skills[key] = int(float(val))
                    except (ValueError, TypeError):
                        pass
            
            available_from = None
            available_until = None
            
            for key in ['available_from', 'available from', 'Available From']:
                if key in row_data:
                    val = row_data[key]
                    if HAS_PANDAS and pd.notna(val):
                        available_from = pd.to_datetime(val)
                    elif val:
                        available_from = datetime.strptime(str(val), '%Y-%m-%d')
            
            for key in ['available_until', 'available until', 'Available Until']:
                if key in row_data:
                    val = row_data[key]
                    if HAS_PANDAS and pd.notna(val):
                        available_until = pd.to_datetime(val)
                    elif val:
                        available_until = datetime.strptime(str(val), '%Y-%m-%d')
            
            self.resource_pool[name] = Person(
                name=name,
                skills=skills,
                available_from=available_from,
                available_until=available_until
            )
        
        print(f"[OK] Loaded {len(self.resource_pool)} persons")
    
    def load_baseline(self, file_path: str) -> None:
        """Load baseline requirements"""
        if HAS_PANDAS:
            df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
        else:
            import csv
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                df = list(reader)
        
        for row in (df.iterrows() if HAS_PANDAS else enumerate(df)):
            row_data = row[1] if HAS_PANDAS else row
            
            month = str(row_data.get('month', row_data.get('Month', row_data.get('month', ''))))
            if not month or month == 'nan':
                continue
            
            count_val = row_data.get('count', row_data.get('Count', row_data.get('required_count', 0)))
            try:
                required_count = int(float(count_val))
            except (ValueError, TypeError):
                required_count = 0
            
            skill_str = row_data.get('skills', row_data.get('Skills', row_data.get('required_skills', '')))
            required_skills = self.parse_skills(skill_str)
            
            self.baseline.append(BaselineRequirement(
                month=month,
                required_count=required_count,
                required_skills=required_skills
            ))
        
        self.baseline.sort(key=lambda x: x.month)
        print(f"[OK] Loaded {len(self.baseline)} months baseline")
    
    def load_original_assignments(self, file_path: str) -> None:
        """Load original assignments"""
        if HAS_PANDAS:
            df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
        else:
            import csv
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                df = list(reader)
        
        for row in (df.iterrows() if HAS_PANDAS else enumerate(df)):
            row_data = row[1] if HAS_PANDAS else row
            
            name = str(row_data.get('name', row_data.get('Name', row_data.get('name', ''))))
            month = str(row_data.get('month', row_data.get('Month', row_data.get('month', ''))))
            role = str(row_data.get('role', row_data.get('Role', row_data.get('role', 'Test'))))
            
            if name and month and name != 'nan' and month != 'nan':
                self.original_assignments.append(Assignment(
                    person_name=name,
                    month=month,
                    role=role,
                    is_original=True
                ))
        
        print(f"[OK] Loaded {len(self.original_assignments)} original assignments")
    
    def load_test_features(self, file_path: str) -> None:
        """Load test features"""
        if HAS_PANDAS:
            df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
        else:
            import csv
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                df = list(reader)
        
        for row in (df.iterrows() if HAS_PANDAS else enumerate(df)):
            row_data = row[1] if HAS_PANDAS else row
            
            feature = str(row_data.get('feature', row_data.get('Feature', row_data.get('feature', ''))))
            if not feature or feature == 'nan':
                continue
            
            skill_str = row_data.get('skills', row_data.get('Skills', row_data.get('required_skills', '')))
            self.test_features[feature] = self.parse_skills(skill_str)
        
        print(f"[OK] Loaded {len(self.test_features)} test features")
    
    def get_assigned_persons(self, month: str) -> Set[str]:
        """Get assigned persons for a month"""
        return {a.person_name for a in self.original_assignments if a.month == month}
    
    def get_available_persons(self, month: str, required_skills: Dict[str, int] = None) -> List[Tuple[str, float]]:
        """Get available persons sorted by skill match"""
        assigned = self.get_assigned_persons(month)
        available = []
        
        month_dt = datetime.strptime(month, "%Y-%m")
        
        for name, person in self.resource_pool.items():
            if name in assigned:
                continue
            
            if person.available_from and month_dt < person.available_from:
                continue
            
            if person.available_until:
                if month_dt.month == 12:
                    month_end = month_dt.replace(year=month_dt.year + 1, month=1, day=1) - timedelta(days=1)
                else:
                    month_end = month_dt.replace(month=month_dt.month + 1, day=1) - timedelta(days=1)
                if month_end > person.available_until:
                    continue
            
            score = person.skill_score(required_skills) if required_skills else 0.5
            available.append((name, score))
        
        available.sort(key=lambda x: x[1], reverse=True)
        return available
    
    def schedule(self, start_date: str, end_date: str) -> List[Assignment]:
        """Execute scheduling"""
        start_dt = datetime.strptime(start_date[:7], "%Y-%m")
        end_dt = datetime.strptime(end_date[:7], "%Y-%m")
        
        months = []
        current = start_dt
        while current <= end_dt:
            months.append(current.strftime("%Y-%m"))
            if current.month == 12:
                current = current.replace(year=current.year + 1, month=1)
            else:
                current = current.replace(month=current.month + 1)
        
        print(f"\nSchedule range: {months[0]} ~ {months[-1]} ({len(months)} months)")
        print("=" * 60)
        
        new_assignments = list(self.original_assignments)
        temp_assigned = list(self.original_assignments)
        
        for month in months:
            baseline_req = None
            for req in self.baseline:
                if req.month == month:
                    baseline_req = req
                    break
            
            if not baseline_req:
                print(f"  {month}: No baseline, skip")
                continue
            
            current_count = len({a.person_name for a in temp_assigned if a.month == month})
            need_count = baseline_req.required_count - current_count
            
            if need_count <= 0:
                print(f"  {month}: OK ({current_count}/{baseline_req.required_count})")
                continue
            
            print(f"  {month}: Need {need_count} more (have {current_count}/{baseline_req.required_count})")
            
            available = self.get_available_persons(month, baseline_req.required_skills)
            
            filled = 0
            for name, score in available:
                if filled >= need_count:
                    break
                
                person = self.resource_pool[name]
                if baseline_req.required_skills:
                    meets = all(
                        person.skills.get(skill, 0) >= level
                        for skill, level in baseline_req.required_skills.items()
                    )
                    if not meets and score < 0.3:
                        continue
                
                new_assignments.append(Assignment(
                    person_name=name,
                    month=month,
                    role="Test",
                    skills_used=list(baseline_req.required_skills.keys()) if baseline_req.required_skills else [],
                    is_original=False
                ))
                
                temp_assigned.append(Assignment(
                    person_name=name,
                    month=month,
                    role="Test",
                    is_original=True
                ))
                
                filled += 1
                print(f"    + {name} (match: {score:.0%})")
            
            if filled < need_count:
                print(f"    Warning: Only filled {filled}/{need_count}")
        
        return new_assignments
    
    def export_schedule(self, assignments: List[Assignment], output_path: str) -> None:
        """Export schedule to file"""
        data = []
        for a in assignments:
            data.append({
                'Month': a.month,
                'Name': a.person_name,
                'Role': a.role,
                'Skills': ', '.join(a.skills_used) if a.skills_used else '-',
                'Original': 'Yes' if a.is_original else 'No'
            })
        
        if HAS_PANDAS:
            df = pd.DataFrame(data)
            df = df.sort_values(['Month', 'Name'])
            if output_path.endswith('.xlsx'):
                df.to_excel(output_path, index=False, engine='openpyxl')
            else:
                df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            import csv
            with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['Month', 'Name', 'Role', 'Skills', 'Original'])
                writer.writeheader()
                for row in sorted(data, key=lambda x: (x['Month'], x['Name'])):
                    writer.writerow(row)
        
        print(f"\n[OK] Exported to: {output_path}")
        print(f"     Total: {len(assignments)} assignments")
    
    def export_monthly_summary(self, assignments: List[Assignment], output_path: str) -> None:
        """Export monthly summary"""
        monthly_data = defaultdict(lambda: {'count': 0, 'names': []})
        
        for a in assignments:
            monthly_data[a.month]['count'] += 1
            monthly_data[a.month]['names'].append(a.person_name)
        
        data = []
        for month in sorted(monthly_data.keys()):
            info = monthly_data[month]
            data.append({
                'Month': month,
                'Count': info['count'],
                'Names': ', '.join(info['names'])
            })
        
        if HAS_PANDAS:
            df = pd.DataFrame(data)
            if output_path.endswith('.xlsx'):
                df.to_excel(output_path, index=False, engine='openpyxl')
            else:
                df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            import csv
            with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['Month', 'Count', 'Names'])
                writer.writeheader()
                for row in data:
                    writer.writerow(row)
        
        print(f"[OK] Monthly summary: {output_path}")


def create_sample_files(output_dir: str = "./sample_data"):
    """Create sample data files"""
    os.makedirs(output_dir, exist_ok=True)
    
    if HAS_PANDAS:
        # Resource pool
        resource_pool = pd.DataFrame([
            {'Name': 'Zhang San', 'Python': 5, 'Java': 3, 'Test': 4, 'Available From': '2024-01-01'},
            {'Name': 'Li Si', 'Python': 3, 'Java': 5, 'Test': 3, 'Available From': '2024-01-01'},
            {'Name': 'Wang Wu', 'Python': 4, 'Java': 2, 'Test': 5, 'Available From': '2024-02-01'},
            {'Name': 'Zhao Liu', 'Python': 2, 'Java': 4, 'Test': 4, 'Available From': '2024-01-01'},
            {'Name': 'Qian Qi', 'Python': 5, 'Java': 5, 'Test': 3, 'Available From': '2024-03-01'},
            {'Name': 'Sun Ba', 'Python': 3, 'Java': 3, 'Test': 5, 'Available From': '2024-01-01'},
            {'Name': 'Zhou Jiu', 'Python': 4, 'Java': 4, 'Test': 4, 'Available From': '2024-01-01'},
            {'Name': 'Wu Shi', 'Python': 5, 'Java': 3, 'Test': 2, 'Available From': '2024-02-01'},
        ])
        resource_pool.to_excel(f"{output_dir}/resource_pool.xlsx", index=False)
        
        # Baseline
        baseline = pd.DataFrame([
            {'Month': '2024-01', 'Count': 3, 'Skills': 'Python:3,Test:3'},
            {'Month': '2024-02', 'Count': 4, 'Skills': 'Python:3,Java:3'},
            {'Month': '2024-03', 'Count': 5, 'Skills': 'Python:4,Test:4'},
            {'Month': '2024-04', 'Count': 4, 'Skills': 'Java:4,Test:3'},
            {'Month': '2024-05', 'Count': 3, 'Skills': 'Python:3,Test:3'},
        ])
        baseline.to_excel(f"{output_dir}/baseline.xlsx", index=False)
        
        # Original assignments
        original = pd.DataFrame([
            {'Month': '2024-01', 'Name': 'Zhang San', 'Role': 'Lead'},
            {'Month': '2024-01', 'Name': 'Li Si', 'Role': 'Dev Support'},
            {'Month': '2024-02', 'Name': 'Zhang San', 'Role': 'Lead'},
            {'Month': '2024-03', 'Name': 'Wang Wu', 'Role': 'Test'},
        ])
        original.to_excel(f"{output_dir}/original_assignments.xlsx", index=False)
        
        # Test features
        features = pd.DataFrame([
            {'Feature': 'API Test', 'Skills': 'Python:4,Test:3'},
            {'Feature': 'UI Auto', 'Skills': 'Python:3,Test:4'},
            {'Feature': 'Perf Test', 'Skills': 'Java:3,Test:4'},
        ])
        features.to_excel(f"{output_dir}/test_features.xlsx", index=False)
    else:
        # CSV fallback
        import csv
        
        with open(f"{output_dir}/resource_pool.csv", 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Name', 'Python', 'Java', 'Test', 'Available From'])
            writer.writerow(['Zhang San', 5, 3, 4, '2024-01-01'])
            writer.writerow(['Li Si', 3, 5, 3, '2024-01-01'])
            writer.writerow(['Wang Wu', 4, 2, 5, '2024-02-01'])
            writer.writerow(['Zhao Liu', 2, 4, 4, '2024-01-01'])
            writer.writerow(['Qian Qi', 5, 5, 3, '2024-03-01'])
            writer.writerow(['Sun Ba', 3, 3, 5, '2024-01-01'])
            writer.writerow(['Zhou Jiu', 4, 4, 4, '2024-01-01'])
            writer.writerow(['Wu Shi', 5, 3, 2, '2024-02-01'])
        
        with open(f"{output_dir}/baseline.csv", 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Month', 'Count', 'Skills'])
            writer.writerow(['2024-01', 3, 'Python:3,Test:3'])
            writer.writerow(['2024-02', 4, 'Python:3,Java:3'])
            writer.writerow(['2024-03', 5, 'Python:4,Test:4'])
            writer.writerow(['2024-04', 4, 'Java:4,Test:3'])
            writer.writerow(['2024-05', 3, 'Python:3,Test:3'])
        
        with open(f"{output_dir}/original_assignments.csv", 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Month', 'Name', 'Role'])
            writer.writerow(['2024-01', 'Zhang San', 'Lead'])
            writer.writerow(['2024-01', 'Li Si', 'Dev Support'])
            writer.writerow(['2024-02', 'Zhang San', 'Lead'])
            writer.writerow(['2024-03', 'Wang Wu', 'Test'])
        
        with open(f"{output_dir}/test_features.csv", 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Feature', 'Skills'])
            writer.writerow(['API Test', 'Python:4,Test:3'])
            writer.writerow(['UI Auto', 'Python:3,Test:4'])
            writer.writerow(['Perf Test', 'Java:3,Test:4'])
    
    print(f"[OK] Created sample files in {output_dir}/")


def main():
    parser = argparse.ArgumentParser(
        description='HR Scheduler - Auto-fill remaining resources',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('--start', type=str, help='Start date (YYYY-MM)')
    parser.add_argument('--end', type=str, help='End date (YYYY-MM)')
    parser.add_argument('--baseline', type=str, help='Baseline file path')
    parser.add_argument('--pool', type=str, help='Resource pool file path')
    parser.add_argument('--skills', type=str, help='Skill matrix file path')
    parser.add_argument('--original', type=str, help='Original assignments file path')
    parser.add_argument('--features', type=str, help='Test features file path')
    parser.add_argument('--output', type=str, default='new_schedule.xlsx', help='Output file path')
    parser.add_argument('--sample', action='store_true', help='Create sample data and run')
    
    args = parser.parse_args()
    
    if args.sample:
        create_sample_files()
        sample_dir = "./sample_data"
        args.start = '2024-01'
        args.end = '2024-05'
        ext = '.xlsx' if HAS_PANDAS else '.csv'
        args.baseline = f"{sample_dir}/baseline{ext}"
        args.pool = f"{sample_dir}/resource_pool{ext}"
        args.skills = f"{sample_dir}/resource_pool{ext}"
        args.original = f"{sample_dir}/original_assignments{ext}"
        args.features = f"{sample_dir}/test_features{ext}"
        args.output = f"{sample_dir}/new_schedule{ext}"
    
    if not args.start or not args.end:
        print("Error: --start and --end are required")
        parser.print_help()
        return
    
    print("=" * 60)
    print("HR Scheduler Started")
    print("=" * 60)
    
    scheduler = HRScheduler()
    
    if args.pool:
        scheduler.load_resource_pool(args.pool)
    if args.baseline:
        scheduler.load_baseline(args.baseline)
    if args.original:
        scheduler.load_original_assignments(args.original)
    if args.features:
        scheduler.load_test_features(args.features)
    
    print("\n" + "=" * 60)
    print("Scheduling...")
    print("=" * 60)
    
    new_assignments = scheduler.schedule(args.start, args.end)
    
    print("\n" + "=" * 60)
    print("Exporting Results")
    print("=" * 60)
    
    scheduler.export_schedule(new_assignments, args.output)
    
    summary_path = args.output.replace('.xlsx', '_summary.xlsx').replace('.csv', '_summary.csv')
    scheduler.export_monthly_summary(new_assignments, summary_path)
    
    print("\n" + "=" * 60)
    print("Done!")
    print("=" * 60)


if __name__ == '__main__':
    main()
