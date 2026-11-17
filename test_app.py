"""
Automated quality checks and test suite for Work Reporter.
Uses dynamic config and isolates test operations so user's real notes are never lost.
"""

import os
os.environ["TESTING"] = "1"
import io
import openpyxl
from datetime import date, timedelta
from app import (
    app, db, DailyNote, TimetableSlot, ScheduleOverride, TimetableImage,
    Sent, AppSetting, parse_notes_fallback, generate_weekly_report_excel, build_week_report_data,
    get_config, get_schedule_for_date, build_schedule_email_html, parse_time_str
)

def run_all_tests():
    print("=== Quality Check 1: App initialization & DB schema ===")
    with app.app_context():
        db.create_all()
        assert DailyNote.query.count() >= 0
        assert TimetableSlot.query.count() >= 0
        assert ScheduleOverride.query.count() >= 0
        assert TimetableImage.query.count() >= 0
        assert Sent.query.count() >= 0
        print("✓ Database models initialized successfully.")

    client = app.test_client()

    print("\n=== Quality Check 2: Authentication flow ===")
    resp = client.get("/")
    assert resp.status_code == 302
    assert "/login" in resp.headers["Location"]
    print("✓ Unauthenticated access safely redirects to /login.")

    resp = client.post("/login", data={"password": "wrongpassword"})
    assert resp.status_code == 200
    assert b"Invalid password" in resp.data
    print("✓ Wrong password rejected with error.")

    active_password = get_config("APP_PASSWORD", "changeme")
    resp = client.post("/login", data={"password": active_password}, follow_redirects=True)
    assert resp.status_code == 200
    assert b"Weekly Notes" in resp.data
    print(f"✓ Logged in successfully using active APP_PASSWORD.")

    print("\n=== Quality Check 3: Notes creation & deletion ===")
    with client.session_transaction() as sess:
        csrf_token = sess.get("csrf_token")
    assert csrf_token is not None

    test_date = "2026-11-09"
    resp = client.post("/notes/save", data={
        "csrf_token": csrf_token,
        "date": test_date,
        "content": "Fixed login bug for 2 hrs, attended sprint planning for 1 hr; reviewed 2 PRs"
    }, follow_redirects=True)
    assert resp.status_code == 200

    with app.app_context():
        note = DailyNote.query.filter_by(date=date.fromisoformat(test_date)).first()
        assert note is not None
        assert "Fixed login bug" in note.content
        print("✓ Daily note persisted in database.")

    # Deleting note by saving empty text
    resp = client.post("/notes/save", data={
        "csrf_token": csrf_token,
        "date": test_date,
        "content": "   "
    }, follow_redirects=True)
    assert resp.status_code == 200
    with app.app_context():
        deleted = DailyNote.query.filter_by(date=date.fromisoformat(test_date)).first()
        assert deleted is None
    print("✓ Saving empty text deletes daily note (shows Empty on dashboard).")

    print("\n=== Quality Check 4: Semester Timetable & Human Input Overrides ===")
    with app.app_context():
        # Add recurring Monday slot
        TimetableSlot.query.filter_by(subject="Algorithms & Data Structures").delete()
        new_slot = TimetableSlot(
            day_of_week=0, # Monday
            start_time="09:00 AM",
            end_time="10:30 AM",
            subject="Algorithms & Data Structures",
            room="Room 304",
            faculty="Dr. Alan"
        )
        db.session.add(new_slot)
        db.session.commit()
        print("✓ Added recurring semester class slot.")

    # Add human input override via POST
    target_tomorrow = date(2026, 11, 16)
    resp = client.post("/schedule/override/add", data={
        "csrf_token": csrf_token,
        "override_date": target_tomorrow.isoformat(),
        "action": "add",
        "start_time": "03:00 PM",
        "end_time": "04:30 PM",
        "subject": "Special Seminar: Cloud Frameworks",
        "room": "Auditorium",
        "notes": "Mandatory attendance"
    }, follow_redirects=True)
    assert resp.status_code == 200

    with app.app_context():
        sched = get_schedule_for_date(target_tomorrow)
        assert sched["has_classes"] is True
        assert len(sched["regular_classes"]) >= 1
        assert len(sched["additional_classes"]) >= 1

        email_html = build_schedule_email_html(sched)
        assert "Algorithms &amp; Data Structures" in email_html or "Algorithms & Data Structures" in email_html
        assert "Special Seminar: Cloud Frameworks" in email_html
        assert "REPLACEMENT CLASS" in email_html or "ADDITIONAL CLASS" in email_html
        print("✓ Verified human input replacement class merged into tomorrow's schedule email!")

    print("\n=== Quality Check 5: Excel Report Generation (.xlsx) ===")
    with app.app_context():
        sample_meta = [
            {"name": "Monday", "date": date(2026, 11, 9), "task_count": 2, "total_hours": 3.0},
            {"name": "Tuesday", "date": date(2026, 11, 10), "task_count": 1, "total_hours": 2.0},
            {"name": "Wednesday", "date": date(2026, 11, 11), "task_count": 0, "total_hours": 0.0},
            {"name": "Thursday", "date": date(2026, 11, 12), "task_count": 0, "total_hours": 0.0},
            {"name": "Friday", "date": date(2026, 11, 13), "task_count": 2, "total_hours": 4.0}
        ]
        sample_tasks = [
            {"day": "Monday", "task": "Fixed auth bug", "category": "Development", "hours": 2, "status": "Completed"},
            {"day": "Monday", "task": "Sprint planning", "category": "Meeting", "hours": 1, "status": "Completed"},
            {"day": "Friday", "task": "Read cloud computing frameworks", "category": "Research", "hours": 4, "status": "Completed"}
        ]
        excel_buf = generate_weekly_report_excel(date(2026, 11, 9), date(2026, 11, 13), sample_tasks, sample_meta)
        wb = openpyxl.load_workbook(excel_buf)
        assert "Weekly Report" in wb.sheetnames
        assert "Summary" in wb.sheetnames
        print("✓ Verified both sheets ('Weekly Report' & 'Summary') in generated Excel.")

    print("\n=== Quality Check 6: Cron Tick endpoint (/cron/tick) ===")
    resp_no_key = client.get("/cron/tick")
    assert resp_no_key.status_code == 403
    print("✓ /cron/tick without key returns 403 Forbidden.")

    cron_key = get_config("CRON_SECRET", "test-cron-secret-xyz")
    resp_valid = client.get(f"/cron/tick?key={cron_key}")
    assert resp_valid.status_code == 200
    json_data = resp_valid.get_json()
    assert json_data["status"] == "ok"
    print(f"✓ /cron/tick with valid key returns 200 OK.")

    print("\n=== Quality Check 7: Excel Timetable Parsing (.xlsx / openpyxl) ===")
    with open("semester_timetable.xlsx", "rb") as f:
        excel_bytes = f.read()

    with app.app_context():
        from app import parse_timetable_excel, build_schedule_email_text
        count, msg = parse_timetable_excel(excel_bytes, filename="semester_timetable.xlsx")
        assert count == 8, f"Expected 8 slots, got {count}: {msg}"
        print(f"✓ Successfully parsed Excel file: {count} slots extracted.")

        # Verify acronym expansion
        dmw_slot = TimetableSlot.query.filter(TimetableSlot.subject.like("%Data Mining and Warehousing%")).first()
        assert dmw_slot is not None
        assert "DMW" in dmw_slot.subject
        assert dmw_slot.room == "MCA B"
        print("✓ Verified acronym expansion: DMW -> Data Mining and Warehousing (DMW) [MCA B]")

        genai_slot = TimetableSlot.query.filter(TimetableSlot.subject.like("%Generative Models%")).first()
        assert genai_slot is not None
        assert "MCA A &B" in genai_slot.room or "MCA A & B" in genai_slot.room
        print("✓ Verified Generative Models slot and MCA A & B section mapping.")

    print("\n=== Quality Check 8: Weekend Manual Entry (Saturday & Sunday) & Simple Email ===")
    test_sat = date(2026, 11, 21) # Saturday
    resp = client.post("/schedule/override/add", data={
        "csrf_token": csrf_token,
        "override_date": test_sat.isoformat(),
        "action": "add",
        "start_time": "10:00 AM",
        "end_time": "12:00 PM",
        "subject": "Data Mining Remedial Session",
        "room": "MCA B / Room 204",
        "notes": "Weekend session"
    }, follow_redirects=True)
    assert resp.status_code == 200

    test_sun = date(2026, 11, 22) # Sunday
    resp = client.post("/schedule/override/add", data={
        "csrf_token": csrf_token,
        "override_date": test_sun.isoformat(),
        "action": "add",
        "start_time": "02:00 PM",
        "end_time": "04:00 PM",
        "subject": "Deep Learning Workshop",
        "room": "Lab 3",
        "notes": "Sunday special guest workshop"
    }, follow_redirects=True)
    assert resp.status_code == 200

    with app.app_context():
        # Check Saturday schedule
        sat_sched = get_schedule_for_date(test_sat)
        assert sat_sched["has_classes"] is True
        assert len(sat_sched["additional_classes"]) == 1

        sat_text = build_schedule_email_text(sat_sched)
        assert "Date: Saturday, 21 Nov 2026" in sat_text
        assert "10:00 AM – 12:00 PM" in sat_text
        assert "Data Mining Remedial Session" in sat_text
        print("✓ Verified Saturday manual entry and simple text email format (Date, Class, Timings).")

        sat_html = build_schedule_email_html(sat_sched)
        assert "Date: Saturday, 21 Nov 2026" in sat_html
        assert "10:00 AM – 12:00 PM" in sat_html
        assert "Data Mining Remedial Session" in sat_html
        print("✓ Verified Saturday HTML email formatting.")

        # Check Sunday schedule
        sun_sched = get_schedule_for_date(test_sun)
        assert sun_sched["has_classes"] is True
        sun_text = build_schedule_email_text(sun_sched)
        assert "Date: Sunday, 22 Nov 2026" in sun_text
        assert "02:00 PM – 04:00 PM" in sun_text
        print("✓ Verified Sunday manual entry and simple email format.")

        # Cleanup test overrides
        ScheduleOverride.query.filter(ScheduleOverride.date.in_([test_sat, test_sun])).delete()
        db.session.commit()

    print("\n=== Quality Check 9: REST API Endpoints (/api/timetable/slots & upload) ===")
    api_slots_resp = client.get("/api/timetable/slots")
    assert api_slots_resp.status_code == 200
    api_data = api_slots_resp.get_json()
    assert api_data["status"] == "success"
    assert api_data["count"] == 8
    print("✓ GET /api/timetable/slots returned 8 slots in JSON.")

    with open("semester_timetable.xlsx", "rb") as f:
        file_data = io.BytesIO(f.read())

    api_upload_resp = client.post(
        "/api/timetable/upload",
        data={"file": (file_data, "timetable.xlsx")},
        content_type="multipart/form-data"
    )
    assert api_upload_resp.status_code == 200
    upload_res = api_upload_resp.get_json()
    assert upload_res["status"] == "success"
    assert upload_res["count"] == 8
    print("✓ POST /api/timetable/upload successfully imported Excel via REST API.")

    print("\n=== Quality Check 10: Robust AM/PM Time Parsing & Database AppSetting ===")
    from datetime import time
    assert parse_time_str("9:00 am") == time(9, 0)
    assert parse_time_str("9 :00 am") == time(9, 0)
    assert parse_time_str("9:00 AM") == time(9, 0)
    assert parse_time_str("09:00 AM") == time(9, 0)
    assert parse_time_str("9:00 pm") == time(21, 0)
    assert parse_time_str("9:00 PM") == time(21, 0)
    assert parse_time_str("18:00") == time(18, 0)
    assert parse_time_str("9 AM") == time(9, 0)
    assert parse_time_str("9 PM") == time(21, 0)
    print("✓ All 12-hour AM/PM and 24-hour time strings parsed accurately.")

    with app.app_context():
        AppSetting.set("SCHEDULE_TIME", "09:00 AM")
        assert get_config("SCHEDULE_TIME") == "09:00 AM"
        assert parse_time_str(get_config("SCHEDULE_TIME")) == time(9, 0)
        print("✓ AppSetting database persistence verified.")

    # Verify /cron/tick response format
    cron_secret = get_config("CRON_SECRET", "test-cron-secret-xyz")
    cron_resp = client.get(f"/cron/tick?key={cron_secret}")
    assert cron_resp.status_code == 200
    cron_data = cron_resp.get_json()
    assert cron_data["status"] == "ok"
    assert "configured_schedule_time" in cron_data or "actions" in cron_data
    print("✓ GET /cron/tick returns detailed diagnostics status.")

    print("\n==========================================")
    print("ALL TESTS PASSED SUCCESSFULLY! 🚀")
    print("==========================================")

if __name__ == "__main__":
    run_all_tests()
