import cProfile
import os
import pstats
import random

import rpyxl
from rpyxl.styles import Alignment, Font, PatternFill

N_ROWS = 50_000
HEADERS = ["Name", "Age", "City", "Job Title", "Email"]

first = [
    "Aarav",
    "Priya",
    "Liam",
    "Sofia",
    "Noah",
    "Meera",
    "Ethan",
    "Isha",
    "Lucas",
    "Zara",
    "Rohan",
    "Emma",
    "Kabir",
    "Olivia",
    "Arjun",
    "Mia",
    "Dev",
    "Ava",
    "Neil",
    "Anika",
]
last = [
    "Sharma",
    "Patel",
    "Smith",
    "Garcia",
    "Khan",
    "Iyer",
    "Brown",
    "Nair",
    "Wilson",
    "Mehta",
    "Kapoor",
    "Lee",
    "Rao",
    "Johnson",
    "Gupta",
    "Martin",
    "Joshi",
    "Clark",
    "Das",
    "Reddy",
]
cities = [
    "Mumbai",
    "Pune",
    "Delhi",
    "Bengaluru",
    "Chennai",
    "London",
    "Toronto",
    "Singapore",
    "Sydney",
    "Dubai",
]
jobs = [
    "Data Analyst",
    "Software Engineer",
    "Product Manager",
    "Designer",
    "Accountant",
    "Marketing Lead",
    "HR Specialist",
    "Sales Executive",
    "Teacher",
    "Nurse",
]

random.seed(42)
rows = []
for i in range(N_ROWS):
    f, l = random.choice(first), random.choice(last)
    rows.append(
        [
            f"{f} {l}",
            random.randint(22, 58),
            random.choice(cities),
            random.choice(jobs),
            f"{f.lower()}.{l.lower()}{i}@example.com",
        ]
    )

print(f"Data ready: {len(rows)} rows.")

head_font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
head_fill = PatternFill("solid", fgColor="1F3A5F")
stripe_fill = PatternFill("solid", fgColor="EAF1FB")
body_font = Font(name="Arial", size=11)
center = Alignment(horizontal="center")


def build_wb():
    wb = rpyxl.Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for c in ws[1]:
        c.font, c.fill = head_font, head_fill
    for i, r in enumerate(rows, start=2):
        stripe = i % 2 == 0
        for j, value in enumerate(r, start=1):
            c = ws.cell(row=i, column=j, value=value)
            c.font = body_font
            if stripe:
                c.fill = stripe_fill
            if j == 2:
                c.alignment = center
    ws.freeze_panes = "A2"
    return wb


print("Profiling BUILD phase...")
pr_build = cProfile.Profile()
pr_build.enable()
wb = build_wb()
pr_build.disable()

print("\n" + "=" * 80)
print("BUILD PHASE PROFILE (Top 20 by tottime):")
print("=" * 80)
ps_build = pstats.Stats(pr_build).sort_stats("tottime")
ps_build.print_stats(20)

print("Profiling SAVE phase...")
save_path = "/tmp/profile_styled_50k.xlsx"
pr_save = cProfile.Profile()
pr_save.enable()
wb.save(save_path)
pr_save.disable()
wb.close()

print("\n" + "=" * 80)
print("SAVE PHASE PROFILE (Top 20 by tottime):")
print("=" * 80)
ps_save = pstats.Stats(pr_save).sort_stats("tottime")
ps_save.print_stats(20)

if os.path.exists(save_path):
    os.remove(save_path)
