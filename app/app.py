import os
import json
import re

import caldav
from datetime import datetime, timedelta

import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def sendEmail(subject, message, type):
  letter = MIMEMultipart("alternative")
  letter["From"] = os.environ.get('SENDER_EMAIL')
  letter["To"] = os.environ.get('RECIEVER_EMAIL')
  letter["Subject"] = subject
  letter.attach(MIMEText(message, type))
  context = ssl.create_default_context()

  try:
    with smtplib.SMTP_SSL(os.environ.get('SMTP_SERVER'), os.environ.get('SMTP_PORT'), context=context) as server:
      server.login(os.environ.get('SENDER_EMAIL'), os.environ.get('SENDER_PASSWORD'))
      server.sendmail(os.environ.get('SENDER_EMAIL'), os.environ.get('RECIEVER_EMAIL'), letter.as_string())
  except:
    print(f"Failed to send email from {os.environ.get('SENDER_EMAIL')} to {os.environ.get('RECIEVER_EMAIL')} trought {os.environ.get('SMTP_SERVER')}:{os.environ.get('SMTP_PORT')}")

try:
  rooms = json.loads(os.environ.get('ROOMS'))
except:
  print("Failed to parse rooms list from environment variable")
  exit(1)

for room in rooms:
  room_url = f"https://calendar.mail.ru/principals/{os.environ.get('DOMAIN')}/{room['name']}"

  try:
    client = caldav.DAVClient(
      url = room_url, 
      auth = (f"{room['name']}@{os.environ.get('DOMAIN')}", room['password'])
    )
    room_calendar = client.principal().calendars()[1]
    events = room_calendar.events()
  except:
    print(f"{room['name']}: Failed to get events")
    continue

  fetched_events = []

  for event in events:
    try:
      formatted_data = re.sub(r"\n\s+", "", event.data)
      partstat = re.findall(rf"ATTENDEE;PARTSTAT=(NEEDS-ACTION);ROLE=REQ-PARTICIPANT:mailto:{room['name']}@{os.environ.get('DOMAIN')}", formatted_data)
      summary = re.findall(r"SUMMARY:(.*)", event.data)[0] if len(re.findall(r"SUMMARY:(.*)", formatted_data)) > 0 else "Untitled"
      organizer = re.findall(r"ORGANIZER;CN=\"(.*)\"", event.data)[0] if len(re.findall(r"ORGANIZER;CN=\"(.*)\"", formatted_data)) > 0 else "Unknown"
    except:
      print(f"{room['name']}: Failed to parse data from event - {event}")
      continue

    if len(partstat) == 1:
      try:
        start = re.findall(r"DTSTART:(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z", formatted_data)
        end = re.findall(r"DTEND:(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z", formatted_data)
        created = re.findall(r"CREATED:(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z", formatted_data)
        start_datetime = datetime(int(start[0][0]), int(start[0][1]), int(start[0][2]), int(start[0][3]), int(start[0][4]), int(start[0][5]))
        end_datetime = datetime(int(end[0][0]), int(end[0][1]), int(end[0][2]), int(end[0][3]), int(end[0][4]), int(start[0][5]))
        created_datetime = datetime(int(created[0][0]), int(created[0][1]), int(created[0][2]), int(created[0][3]), int(created[0][4]), int(created[0][5]))
      except:
        print(f"{room['name']}: Failed to parse data - {summary} by {organizer}")
        continue

      if (start_datetime >= datetime.now() and end_datetime <= datetime.now() + timedelta(days=int(os.environ.get('DAYS_NUMBER')))):
        fetched_events.append({ 
          'url': event.url,
          'name': summary,
          'start': start_datetime, 
          'end': end_datetime,
          'created': created_datetime,
          'organizer': organizer
        })

    fetched_events.sort(key=lambda event: event['created'])

  for first in fetched_events:
    second_events = fetched_events[:]
    second_events.pop(second_events.index(first))
    for second in second_events:
      latest_start = max(first['start'], second['start'])
      earliest_end = min(first['end'], second['end'])
      overlap_days = max(0, (earliest_end - latest_start).days + 1)
      overlap_seconds = max(0, (earliest_end - latest_start).seconds)
      if overlap_days > 0 and overlap_seconds > 0:
        start_time = second['start'].strftime('%d-%b-%Y (%H:%M:%S)')
        try:
          room_calendar.event_by_url(second['url']).decline_invite(calendar=room_calendar)
        except:
          print(f"{room['name']}: Failed to decline event - [{start_time}] {second['name']} by {second['organizer']}")
          subject = f"Проблема с бронированием в {room['name']}"
          message = f"""\
            <html>
              <body>
                <ul>
                  <li>Переговорная: <b>{room['name']}</b></li>
                  <li>Встреча: <b>{second['name']}</b></li>
                  <li>Время начала: <b>{start_time}</b></li>
                  <li>Организатор: <b>{second['organizer']}</b></li>
                  <li>Ошибка: <span style="color:red;"><b>Невозможно отклонить встречу</b></span></li>
                </ul>
                <p style="color:#28b463;"><b>Требуется отклонить встречу вручную.</b></p>
              </body>
            </html>
          """
          sendEmail(subject, message, "html")
        else:
          print(f"{room['name']}: Declined - [{start_time}] {second['name']} by {second['organizer']} (Overlap with {first['name']} - {overlap_seconds/60} minutes)")
          fetched_events.pop(fetched_events.index(second))
          
  if len(fetched_events):
    for event in fetched_events:
      start_time = event['start'].strftime('%d-%b-%Y (%H:%M:%S)')
      try:
        room_calendar.event_by_url(event['url']).accept_invite(calendar=room_calendar)
      except:
        print(f"{room['name']}: Failed to accept event - [{start_time}] {event['name']} by {event['organizer']}")
        subject = f"Проблема с бронированием в {room['name']}"
        message = f"""\
          <html>
            <body>
              <ul>
                <li>Переговорная: <b>{room['name']}</b></li>
                <li>Встреча: <b>{event['name']}</b></li>
                <li>Время начала: <b>{start_time}</b></li>
                <li>Организатор: <b>{event['organizer']}</b></li>
                <li>Ошибка: <span style="color:red;"><b>Невозможно принять встречу</b></span></li>
              </ul>
              <p style="color:#28b463;"><b>Требуется принять встречу вручную.</b></p>
            </body>
          </html>
        """
        sendEmail(subject, message, "html")
      else:
        print(f"{room['name']}: Accepted - [{start_time}] {event['name']} by {event['organizer']}")
  else:
    print(f"Nothing to handle in {room['name']}")