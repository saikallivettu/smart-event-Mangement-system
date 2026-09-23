

import os
import json
import urllib.request
import urllib.parse
from datetime import datetime
import openpyxl

try:
    from milestone_1 import (
        events, venues, resources, EXCEL_FILE,
        save_m1_to_excel, find_event, find_venue, find_resource
    )
    from milestone_2 import (
        vendors, save_m2_to_excel, find_vendor
    )
except ImportError:
    print("Error: Ensure 'milestone_1.py' and 'milestone_2.py' are in the same directory!")
    exit(1)


expenses = {}
sponsors = {}
approvals = []
audit_log = []
payments_log = []

expense_number = 101
next_sponsor_id = 1
next_approval_id = 1
payment_counter = 5001


def money(amount):
    return f"Rs. {amount:,.2f}"


def log_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log_action(action):
    audit_log.append({"time": log_time(), "action": action})
    save_m3_to_excel()



def total_expenses(event_id):
    total = 0.0
    for expense in expenses.values():
        if expense["event_id"] == event_id and expense["status"] != "Rejected":
            total += float(expense["amount"])
    return total


def total_committed(event_id):
    total = 0.0
    for expense in expenses.values():
        if expense["event_id"] == event_id and expense["status"] == "Pending":
            total += float(expense["amount"])
    return total


def total_sponsorship(event_id):
    total = 0.0
    for sponsor in sponsors.values():
        if sponsor["event_id"] == event_id and sponsor["status"] == "Confirmed":
            total += float(sponsor["amount"])
    return total



def save_m3_to_excel():
    save_m2_to_excel()

    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
    else:
        wb = openpyxl.Workbook()

    
    if "Expenses" in wb.sheetnames:
        ws_ex = wb["Expenses"]
        ws_ex.delete_rows(1, ws_ex.max_row)
    else:
        ws_ex = wb.create_sheet("Expenses")

    ws_ex.append(["Expense ID", "Event ID", "Category", "Description", "Amount", "Status"])
    for ex in expenses.values():
        ws_ex.append([ex["id"], ex["event_id"], ex["category"], ex["description"], ex["amount"], ex["status"]])

    if "Sponsors" in wb.sheetnames:
        ws_s = wb["Sponsors"]
        ws_s.delete_rows(1, ws_s.max_row)
    else:
        ws_s = wb.create_sheet("Sponsors")

    ws_s.append(["Sponsor ID", "Event ID", "Sponsor Name", "Amount", "Status", "Date"])
    for s in sponsors.values():
        ws_s.append([s["id"], s["event_id"], s["name"], s["amount"], s["status"], s["date"]])

    if "Approvals" in wb.sheetnames:
        ws_a = wb["Approvals"]
        ws_a.delete_rows(1, ws_a.max_row)
    else:
        ws_a = wb.create_sheet("Approvals")

    ws_a.append(["Approval ID", "Type", "Item ID", "Event ID", "Description", "Status"])
    for a in approvals:
        ws_a.append([a["id"], a["type"], a["item_id"], a["event_id"], a["description"], a["status"]])


    if "Audit_Logs" in wb.sheetnames:
        ws_log = wb["Audit_Logs"]
        ws_log.delete_rows(1, ws_log.max_row)
    else:
        ws_log = wb.create_sheet("Audit_Logs")

    ws_log.append(["Timestamp", "Action Log"])
    for entry in audit_log:
        ws_log.append([entry["time"], entry["action"]])

    wb.save(EXCEL_FILE)


def load_m3_from_excel():
    global expense_number, next_sponsor_id, next_approval_id
    if not os.path.exists(EXCEL_FILE):
        return

    wb = openpyxl.load_workbook(EXCEL_FILE)

    if "Expenses" in wb.sheetnames:
        ws_ex = wb["Expenses"]
        max_ex_id = 100
        for row in list(ws_ex.iter_rows(values_only=True))[1:]:
            if row and row[0] is not None:
                ex_id = int(row[0])
                expenses[ex_id] = {
                    "id": ex_id,
                    "event_id": int(row[1]),
                    "category": str(row[2]),
                    "description": str(row[3]),
                    "amount": float(row[4]),
                    "status": str(row[5])
                }
                if ex_id > max_ex_id:
                    max_ex_id = ex_id
        expense_number = max_ex_id + 1

    if "Sponsors" in wb.sheetnames:
        ws_s = wb["Sponsors"]
        max_s_id = 0
        for row in list(ws_s.iter_rows(values_only=True))[1:]:
            if row and row[0] is not None:
                s_id = int(row[0])
                sponsors[s_id] = {
                    "id": s_id,
                    "event_id": int(row[1]),
                    "name": str(row[2]),
                    "amount": float(row[3]),
                    "status": str(row[4]),
                    "date": str(row[5])
                }
                if s_id > max_s_id:
                    max_s_id = s_id
        next_sponsor_id = max_s_id + 1

    if "Approvals" in wb.sheetnames:
        ws_a = wb["Approvals"]
        max_a_id = 0
        for row in list(ws_a.iter_rows(values_only=True))[1:]:
            if row and row[0] is not None:
                a_id = int(row[0])
                approvals.append({
                    "id": a_id,
                    "type": str(row[1]),
                    "item_id": int(row[2]),
                    "event_id": int(row[3]),
                    "description": str(row[4]),
                    "status": str(row[5])
                })
                if a_id > max_a_id:
                    max_a_id = a_id
        next_approval_id = max_a_id + 1

    if "Audit_Logs" in wb.sheetnames:
        ws_log = wb["Audit_Logs"]
        for row in list(ws_log.iter_rows(values_only=True))[1:]:
            if row and row[0] is not None:
                audit_log.append({"time": str(row[0]), "action": str(row[1])})


load_m3_from_excel()


def add_sponsor(event_id, name, amount):
    global next_sponsor_id, next_approval_id

    for s in sponsors.values():
        if s["event_id"] == event_id and s["name"].lower() == name.lower():
            print(f"Note: Sponsor {name} has already been added for this event.")
            return

    sponsors[next_sponsor_id] = {
        "id": next_sponsor_id,
        "event_id": event_id,
        "name": name,
        "amount": amount,
        "status": "Pending",
        "date": log_time()
    }

    approvals.append({
        "id": next_approval_id,
        "type": "Sponsor",
        "item_id": next_sponsor_id,
        "event_id": event_id,
        "description": f"{name} offering {money(amount)}",
        "status": "Pending"
    })

    print(f"Sponsor request recorded: {name} pledged {money(amount)} and is waiting for approval.")
    log_action(f"Sponsor added: {name}")

    next_sponsor_id += 1
    next_approval_id += 1


def add_sponsor_interactive():
    if not events:
        print("No events available.")
        return

    print("\nAvailable Events:")
    for e in events:
        print(f"  ID: {e.id} | Name: {e.name}")

    try:
        event_id = int(input("Enter Event ID: "))
        amount = float(input("Enter Pledged Amount (Rs.): "))
    except ValueError:
        print("Invalid input!")
        return

    name = input("Enter Sponsor Name: ").strip()
    if not name:
        print("Sponsor name cannot be empty.")
        return

    add_sponsor(event_id, name, amount)


def view_sponsors():
    print("\nCurrent sponsors for this event:")
    if not sponsors:
        print("No sponsors have been added yet.")
        return
    for s in sponsors.values():
        print(f"- {s['name']} pledged {money(s['amount'])} ({s['status']} on {s['date']})")


def view_approvals():
    print("\nApproval requests:")
    if not approvals:
        print("No approvals are pending.")
        return
    for a in approvals:
        print(f"- Request {a['id']} for Event {a['event_id']}: {a['description']} ({a['status']})")


def process_approval(approval_id, decision):
    approval = next((a for a in approvals if a["id"] == approval_id), None)
    if not approval:
        print("No approval found with that ID.")
        return
    if approval["status"] != "Pending":
        print("This request has already been processed.")
        return

    if decision.upper() == "APPROVE":
        approval["status"] = "Approved"
        sponsors[approval["item_id"]]["status"] = "Confirmed"
        print(f"Approval complete: {approval['description']} has been confirmed.")
        log_action(f"Approval #{approval_id} approved")
    elif decision.upper() == "REJECT":
        approval["status"] = "Rejected"
        sponsors[approval["item_id"]]["status"] = "Rejected"
        print(f"Approval complete: {approval['description']} has been rejected.")
        log_action(f"Approval #{approval_id} rejected")
    else:
        print("Invalid choice. Please use APPROVE or REJECT.")


def process_approval_interactive():
    view_approvals()
    if not approvals:
        return

    try:
        approval_id = int(input("\nEnter Approval Request ID to process: "))
    except ValueError:
        print("Invalid Request ID.")
        return

    decision = input("Enter decision (APPROVE / REJECT): ").strip()
    process_approval(approval_id, decision)


def sponsorship_report(event_id):
    confirmed = sum(s["amount"] for s in sponsors.values()
                    if s["event_id"] == event_id and s["status"] == "Confirmed")
    print(f"\nSummary for Event {event_id}:")
    print(f"Total confirmed sponsorship: {money(confirmed)}")


def sponsorship_report_interactive():
    try:
        event_id = int(input("Enter Event ID for Sponsorship Summary: "))
    except ValueError:
        print("Invalid Event ID.")
        return
    sponsorship_report(event_id)


def search_sponsors(keyword):
    results = [s for s in sponsors.values() if keyword.lower() in s["name"].lower()]
    print(f"\nSearch results for '{keyword}':")
    if not results:
        print("No sponsors matched your search.")
    else:
        for s in results:
            print(f"- {s['name']} pledged {money(s['amount'])} ({s['status']})")


def search_sponsors_interactive():
    keyword = input("\nEnter Sponsor Name to search: ").strip()
    if keyword:
        search_sponsors(keyword)


def set_budget():
    print("\n================ SET EVENT BUDGET ================")
    if not events:
        print("No events available.")
        return

    for e in events:
        print(f"ID: {e.id} | Name: {e.name}")

    try:
        event_id = int(input("Enter Event ID: "))
    except ValueError:
        print("Invalid Event ID.")
        return

    event = find_event(event_id)
    if event is None:
        print("Event not found.")
        return

    try:
        budget = float(input("Enter allocated budget: "))
        if budget < 0:
            raise ValueError
    except ValueError:
        print("Enter a valid non-negative budget.")
        return

    event.budget = budget
    save_m3_to_excel()
    log_action(f"Budget set for Event '{event.name}': {money(budget)}")
    print("Budget saved successfully!")
    print("Event :", event.name)
    print("Budget:", money(budget))


def add_expense():
    global expense_number
    print("\n================ ADD EXPENSE ================")
    if not events:
        print("No events available.")
        return

    for e in events:
        print(f"ID: {e.id} | Name: {e.name}")

    try:
        event_id = int(input("Enter Event ID: "))
    except ValueError:
        print("Invalid Event ID.")
        return

    event = find_event(event_id)
    if event is None:
        print("Event not found.")
        return

    category = input("Enter expense category: ").strip()
    description = input("Enter expense description: ").strip()
    if not category or not description:
        print("Category and description cannot be empty.")
        return

    try:
        amount = float(input("Enter expense amount: "))
        if amount <= 0:
            raise ValueError
    except ValueError:
        print("Enter a valid positive amount.")
        return

    expense_id = expense_number
    expense_number += 1

    expense = {
        "id": expense_id,
        "event_id": event_id,
        "category": category,
        "description": description,
        "amount": amount,
        "status": "Approved"
    }

    expenses[expense_id] = expense
    save_m3_to_excel()
    log_action(f"Expense added: {money(amount)} for Event ID {event_id} ({category})")

    print("Expense added successfully!")
    print("Expense ID:", expense_id)
    print("Amount    :", money(amount))


def view_expenses():
    print("\n================ EXPENSES ================")
    if not expenses:
        print("No expenses found.")
        return

    for expense in expenses.values():
        event = find_event(expense["event_id"])
        event_name = event.name if event else "Unknown Event"
        print(f"\nExpense ID : {expense['id']}")
        print(f"Event      : {event_name}")
        print(f"Category   : {expense['category']}")
        print(f"Description: {expense['description']}")
        print(f"Amount     : {money(expense['amount'])}")
        print(f"Status     : {expense['status']}")


def budget_report():
    print("\n================ BUDGET REPORT ================")
    if not events:
        print("No events available.")
        return

    for event in events:
        budget = float(getattr(event, "budget", 0.0))
        spent = total_expenses(event.id)
        committed = total_committed(event.id)
        sponsorship = total_sponsorship(event.id)
        available = budget + sponsorship
        remaining = available - spent

        if budget > 0:
            utilization = (spent / budget) * 100
        else:
            utilization = 0

        print(f"\nEvent             : {event.name}")
        print(f"Budget            : {money(budget)}")
        print(f"Sponsorship       : {money(sponsorship)}")
        print(f"Total Available   : {money(available)}")
        print(f"Actual Spending   : {money(spent)}")
        print(f"Committed Cost    : {money(committed)}")
        print(f"Remaining Budget  : {money(remaining)}")
        print(f"Budget Utilization: {utilization:.2f}%")

        if budget == 0:
            print("Status: Budget not allocated")
        elif remaining < 0:
            print("Status: BUDGET EXCEEDED")
        elif utilization >= 80:
            print("Status: WARNING - 80% BUDGET USED")
        else:
            print("Status: WITHIN BUDGET")



#1. Weather API
def weather_for_event():
    try:
        event_id = int(input("Enter Event ID: "))
    except ValueError:
        return

    event = find_event(event_id)
    if not event:
        print("Event Not Found.")
        return

    city = input("Enter City name: ").strip()
    if not city:
        return

    try:
        geo_url = f"https://geocoding-api.open-meteo.com/v1/search?name={urllib.parse.quote(city)}&count=1"
        with urllib.request.urlopen(geo_url, timeout=8) as res:
            geo_data = json.loads(res.read().decode())

        if not geo_data.get("results"):
            print(f"City '{city}' not found.")
            return

        place = geo_data["results"][0]
        weather_url = (
            f"https://api.open-meteo.com/v1/forecast"
            f"?latitude={place['latitude']}&longitude={place['longitude']}"
            f"&daily=temperature_2m_max,precipitation_probability_max"
            f"&timezone=auto&start_date={event.date}&end_date={event.date}"
        )
        with urllib.request.urlopen(weather_url, timeout=8) as res:
            w_data = json.loads(res.read().decode())

        temp = w_data["daily"]["temperature_2m_max"][0]
        rain = w_data["daily"]["precipitation_probability_max"][0]
        print(f"\nWeather for '{event.name}' in {place['name']} on {event.date}:")
        print(f"  High Temp  : {temp}°C")
        print(f"  Rain Chance: {rain}%")
    except Exception as e:
        print(f"Weather lookup failed: {e}")


# 2. Maps & Location API
def maps_location_lookup():
    venue_name = input("\nEnter Venue / Location Name: ").strip()
    if not venue_name:
        print("Venue name cannot be empty.")
        return

    encoded_name = urllib.parse.quote(venue_name)
    maps_url = f"https://www.google.com/maps/search/?api=1&query={encoded_name}"

    print(f"\n[Maps & Location API]")
    print(f"  Location Query : {venue_name}")
    print(f"  Directions Link: {maps_url}")
    log_action(f"Maps lookup performed for '{venue_name}'")


# 3. Email API
def email_api(to_email, subject, body):
    if not isinstance(to_email, str) or "@" not in to_email:
        print(f"\n[Email API 400 Bad Request]: Invalid email address '{to_email}'")
        return 400

    print(f"\n[Email API 201 Created] Email queued for '{to_email}'")
    print(f"  Subject: {subject}")
    log_action(f"Email sent to {to_email}: '{subject}'")
    return 201


def trigger_email_confirmation():
    email = input("\nEnter Participant Email: ").strip()
    subject = input("Enter Email Subject: ").strip()
    body = input("Enter Email Body: ").strip()
    email_api(email, subject, body)


# 4. SMS API
def sms_api(phone_number, message):
    clean_phone = "".join(filter(str.isdigit, str(phone_number)))
    if len(clean_phone) < 10:
        print(f"\n[SMS API 400 Error]: Invalid phone number '{phone_number}'")
        return 400

    print(f"\n[SMS API 200 OK] Urgent SMS Dispatched to +91-{clean_phone}:")
    print(f"  SMS Text: {message}")
    log_action(f"Urgent SMS sent to {clean_phone}")
    return 200


def send_urgent_sms_update():
    phone = input("\nEnter Phone Number: ").strip()
    msg = input("Enter Urgent Message: ").strip()
    sms_api(phone, msg)


# 5. Payment API
def process_payment_api():
    global payment_counter
    if not events:
        print("No events available.")
        return

    print("\nAvailable Events:")
    for e in events:
        print(f"  ID: {e.id} | Name: {e.name}")

    try:
        event_id = int(input("\nEnter Event ID: "))
        amount = float(input("Enter Registration Fee Amount (Rs.): "))
    except ValueError:
        print("Invalid numerical input.")
        return

    event = find_event(event_id)
    if not event:
        print("Event Not Found.")
        return

    payer_name = input("Enter Payer Name: ").strip()
    txn_id = f"TXN-{payment_counter}"
    payment_counter += 1

    record = {
        "txn_id": txn_id,
        "event_id": event_id,
        "event_name": event.name,
        "payer": payer_name,
        "amount": amount,
        "status": "SUCCESS",
        "time": log_time()
    }
    payments_log.append(record)

    print(f"\n[Payment API 200 OK] Transaction Processed Successfully!")
    print(f"  Txn ID   : {txn_id}")
    print(f"  Payer    : {payer_name}")
    print(f"  Amount   : {money(amount)}")
    print(f"  Status   : SUCCESS")

    log_action(f"Payment received: {money(amount)} from {payer_name} [{txn_id}]")


# 6. Calendar API
def calendar_api_invite():
    if not events:
        print("No events available.")
        return

    print("\nAvailable Events:")
    for e in events:
        print(f"  ID: {e.id} | Name: {e.name} | Date: {e.date}")

    try:
        event_id = int(input("\nEnter Event ID: "))
    except ValueError:
        print("Invalid Event ID.")
        return

    event = find_event(event_id)
    if not event:
        print("Event Not Found.")
        return

    clean_date = event.date.replace("-", "")
    gcal_url = (
        f"https://calendar.google.com/calendar/render?action=TEMPLATE"
        f"&text={urllib.parse.quote(event.name)}"
        f"&dates={clean_date}T100000Z/{clean_date}T170000Z"
        f"&details={urllib.parse.quote('Event registration confirmed via Smart Event Platform')}"
    )

    ics_filename = f"event_{event.id}.ics"
    ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Smart Event Platform//EN
BEGIN:VEVENT
SUMMARY:{event.name}
DTSTART:{clean_date}T100000Z
DTEND:{clean_date}T170000Z
DESCRIPTION:Event registration confirmed.
END:VEVENT
END:VCALENDAR"""

    with open(ics_filename, "w") as f:
        f.write(ics_content)

    print(f"\n[Calendar API] Event Added to Calendar System:")
    print(f"  Google Calendar Link: {gcal_url}")
    print(f"  Saved Calendar File : '{ics_filename}'")
    log_action(f"Calendar invite generated for '{event.name}'")


def view_audit_log():
    print("\nActivity log:")
    if not audit_log:
        print("No actions have been recorded yet.")
        return
    for entry in audit_log:
        print(f"- {entry['time']}: {entry['action']}")


def main():
    while True:
        print("""
=====================================================
 SMART EVENT PLANNING PLATFORM - MILESTONE 3
=====================================================
1. Add Sponsor Request
2. View Current Sponsors
3. View Pending Approval Queue
4. Process Approval / Rejection Decision
5. Sponsorship Financial Summary Report
6. Search Sponsors
7. Set Event Budget
8. Add Expense
9. View All Expenses
10. View Comprehensive Budget Report
11. Weather API (Live Open-Meteo Forecast)
12. Maps & Location API (Venue Query & Directions)
13. Email API (Trigger Confirmation Email)
14. SMS API (Send Urgent Alert SMS)
15. Payment API (Process Registration Fees)
16. Calendar API (Google Calendar & .ics Invites)
17. View Activity / Audit Log
18. Exit
=====================================================
""")
        choice = input("Choice (1-18): ").strip()
        if choice == "1": add_sponsor_interactive()
        elif choice == "2": view_sponsors()
        elif choice == "3": view_approvals()
        elif choice == "4": process_approval_interactive()
        elif choice == "5": sponsorship_report_interactive()
        elif choice == "6": search_sponsors_interactive()
        elif choice == "7": set_budget()
        elif choice == "8": add_expense()
        elif choice == "9": view_expenses()
        elif choice == "10": budget_report()
        elif choice == "11": weather_for_event()
        elif choice == "12": maps_location_lookup()
        elif choice == "13": trigger_email_confirmation()
        elif choice == "14": send_urgent_sms_update()
        elif choice == "15": process_payment_api()
        elif choice == "16": calendar_api_invite()
        elif choice == "17": view_audit_log()
        elif choice == "18":
            print("Exiting Milestone 3 System. Thank You!")
            break
        else:
            print("Invalid Choice! Please enter a number between 1 and 18.")


if __name__ == "__main__":
    main()