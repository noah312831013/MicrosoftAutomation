# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import json
import requests
from datetime import datetime
from django.utils.dateparse import parse_datetime
from django.db import models
from typing import TYPE_CHECKING, List, Dict, Any, Optional
from graph_tutorial.settings import ALLOWED_HOSTS
import aiohttp
import asyncio
GRAPH_URL = 'https://graph.microsoft.com/v1.0'

if TYPE_CHECKING:
    from .models import AutoScheduleMeeting

def get_user(token):
    # Send GET to /me
    user = requests.get(
        f'{GRAPH_URL}/me',
        headers={
          'Authorization': f'Bearer {token}'
        },
        params={
          '$select': 'displayName,mail,mailboxSettings,userPrincipalName'
        })
    # Return the JSON result
    return user.json()

def get_users(token, query=None):
    if not query:
        return []
    query = query.strip()
    if not query:
        raise ValueError("Query parameter is required")

    filter_query = f"startswith(displayName,'{query}') or startswith(mail,'{query}')"
    endpoint = f"{GRAPH_URL}/users?$filter={filter_query}&$select=displayName,mail,userPrincipalName"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    try:
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        data = response.json()

        if "value" not in data:
            raise ValueError("Invalid response from Graph API")

        contacts = []
        for user in data["value"]:
            email = user.get("mail") or user.get("userPrincipalName")
            if email:
                contacts.append({
                    "name": user.get("displayName", ""),
                    "email": email
                })

        return contacts

    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Graph API request failed: {e}")



def get_calendar_events(token, start, end, timezone):
    # Set headers
    headers = {
        'Authorization': f'Bearer {token}',
        'Prefer': f'outlook.timezone="{timezone}"'
    }

    # Configure query parameters to
    # modify the results
    query_params = {
        'startDateTime': start,
        'endDateTime': end,
        '$select': 'subject,organizer,start,end',
        '$orderby': 'start/dateTime',
        '$top': '50'
    }

    # Send GET to /me/events
    events = requests.get(f'{GRAPH_URL}/me/calendarview',
        headers=headers,
        params=query_params)

    # Return the JSON result
    return events.json()

def get_meeting_times_slots(token: str, meeting: 'AutoScheduleMeeting', timezone: str = 'UTC') -> List[Dict[str, Any]]:
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Prefer': f'outlook.timezone="{timezone}"'
    }
    attendees_list = json.loads(meeting.attendees)  # 將 JSON 字符串轉換為列表
    attendees_list.append(meeting.host_email)

    body = {
        "attendees": [
            {
                "emailAddress": { "address": email },
                "type": "Required"
            } for email in attendees_list  # 確保將 JSON 字符串轉換為列表
        ],
        "timeConstraint": {
            "timeslots": [
                {
                    "start": {
                        "dateTime": meeting.start_time.replace(tzinfo=None).isoformat(),
                        "timeZone": timezone
                    },
                    "end": {
                        "dateTime": meeting.end_time.replace(tzinfo=None).isoformat(),
                        "timeZone": timezone
                    }
                }
            ]
        },
        "meetingDuration": f"PT{meeting.duration}M"
    }

    response = requests.post(f'{GRAPH_URL}/me/findMeetingTimes', headers=headers, json=body)

    if response.status_code != 200:
        raise Exception(f"Microsoft Graph API Error: {response.status_code} {response.text}")

    data = response.json()

    # Check if there are no available time slots
    if not data.get("meetingTimeSuggestions"):
        raise Exception("No available meeting time slots found in the response.")

    meeting_times = []
    for suggestion in data.get("meetingTimeSuggestions", []):
        slot = suggestion["meetingTimeSlot"]
        start_dt = parse_datetime(slot["start"]["dateTime"]).isoformat()
        end_dt = parse_datetime(slot["end"]["dateTime"]).isoformat()
        meeting_times.append({
            "confidence": suggestion["confidence"],
            "attendeeAvailability": suggestion["attendeeAvailability"],
            "start": start_dt,
            "end": end_dt
        })


    return meeting_times

def get_user_info(token, email):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }
    response = requests.get(f'{GRAPH_URL}/users/{email}', headers = headers)
    user_data = response.json()
    return user_data

def get_all_chats(token):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    chats = []
    url = f"{GRAPH_URL}/me/chats"
    
    while url:
        res = requests.get(url, headers=headers)
        res.raise_for_status()
        data = res.json()
        chats.extend(data.get('value', []))
        url = data.get('@odata.nextLink')

    return chats

# 一次找全部
def get_chat_ids(token, user_ids):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }
    chats = get_all_chats(token)
    if not chats:
        raise Exception(f"Failed to get chats")
    
    chat_ids = []

    for user_id in user_ids:
        matched_chat_id = None

        for chat in chats:
            if chat.get("chatType") != "oneOnOne":
                continue

            # Get chat members
            chat_id = chat.get("id")
            members_resp = requests.get(f"{GRAPH_URL}/chats/{chat_id}/members", headers=headers)
            if members_resp.status_code != 200:
                continue  # skip if can't fetch members

            members = members_resp.json().get("value", [])

            # Find the "other" member's tenantId
            for member in members:
                if member.get("userId") == user_id:
                    matched_chat_id = chat_id
                    break

            if matched_chat_id:
                break  # found matching chat, no need to check more chats

        chat_ids.append(matched_chat_id)

    return chat_ids

def create_card_payload(subject, start_time, end_time, tenant_id, uuid, base_response_url='http:/localhost/webhook/response/'):
    card = {
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {
                "type": "TextBlock",
                "text": f"📢 會議邀請: {subject}",
                "weight": "Bolder",
                "size": "Medium"
            },
            {
                "type": "TextBlock",
                "text": f"🕒 時間: {start_time} ~ {end_time}"
            }
        ],
        "actions": [
            {
                "type": "Action.OpenUrl",
                "title": "✅ 參加",
                "url": f"{base_response_url}?tenantId={tenant_id}&uuid={str(uuid)}&response=accepted"
            },
            {
                "type": "Action.OpenUrl",
                "title": "❌ 不參加",
                "url": f"{base_response_url}?tenantId={tenant_id}&uuid={str(uuid)}&response=declined"
            }
        ]
    }

    card_payload = {
        "body": {
            "contentType": "html",
            "content": "This message was sent automatically by the Microsoft Automation Tool. <attachment id=\"1\"></attachment>"
        },
        "attachments": [
            {
                "id": "1",
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": json.dumps(card)
            }
        ]
    }

    return card_payload

# # 非同步通知單一與會者
# async def inform_attendee(session, token, email, data, meeting, msg):
#     chat_id = data.get('chat_id')
#     tenant_id = data.get('tenant_id')

#     if not chat_id:
#         print(f"[⚠️] No chat_id for {email}, skipping")
#         return

#     if not msg:
#         card_payload = create_card_payload(
#             subject=meeting.title,
#             start_time=meeting.get_candidate_time()['start'],
#             end_time=meeting.get_candidate_time()['end'],
#             tenant_id=tenant_id,
#             uuid=meeting.uuid,
#             base_response_url="https://c84b-60-248-185-20.ngrok-free.app/webhook/response/"
#         )
#     else:
#         card_payload = {
#             "body": {
#                 "contentType": "html",
#                 "content": msg
#             },
#             "attachments": []
#         }
#     url = f"{GRAPH_URL}/chats/{chat_id}/messages"
#     headers = {
#         'Authorization': f'Bearer {token}',
#         'Content-Type': 'application/json',
#     }

#     async with session.post(url, headers=headers, json=card_payload) as response:
#         if response.status >= 300:
#             text = await response.text()
#             print(f"[❌] Failed to send card to {email} (chat_id: {chat_id})")
#             print(f"Response: {response.status} - {text}")
#         else:
#             print(f"[✅] Message sent to {email}")

# # 非同步通知所有與會者
# async def inform_attendees(token, meeting: 'AutoScheduleMeeting', msg=None):
#     attendee_responses = meeting.get_attendee_responses()

#     async with aiohttp.ClientSession() as session:
#         tasks = [
#             inform_attendee(session, token, email, data, meeting, msg)
#             for email, data in attendee_responses.items()
#         ]
#         await asyncio.gather(*tasks)

def inform_attendees(token, meeting: 'AutoScheduleMeeting'):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }

    attendee_responses = meeting.get_attendee_responses()

    for email, data in attendee_responses.items():
        chat_id = data.get('chat_id')
        tenant_id = data.get('tenant_id')

        if chat_id:
            card_payload = create_card_payload(
                subject=meeting.title,
                start_time=meeting.get_candidate_time()['start'],
                end_time=meeting.get_candidate_time()['end'],
                tenant_id=tenant_id,
                uuid = meeting.uuid,
                base_response_url="https://c84b-60-248-185-20.ngrok-free.app/webhook/response/"
            )

            url = f"{GRAPH_URL}/chats/{chat_id}/messages"

            response = requests.post(url, headers=headers, json=card_payload)

            if response.status_code >= 300:
                print(f"[❌] Failed to send card to {email} (chat_id: {chat_id})")
                print(f"Response: {response.status_code} - {response.text}")
            else:
                print(f"[✅] Card sent to {email}")
        else:
            print(f"[⚠️] No chat_id for {email}, skipping")



def create_event(token, subject, start, end, attendees=None, body=None, timezone='UTC'):
    # Create an event object
    # https://docs.microsoft.com/graph/api/resources/event?view=graph-rest-1.0
    new_event = {
        'subject': subject,
        'start': {
            'dateTime': start,
            'timeZone': timezone
        },
        'end': {
            'dateTime': end,
            'timeZone': timezone
        },
        'location': {
            'displayName': "Teams 線上會議",
        },
        'isOnlineMeeting': True,
        'onlineMeetingProvider': "teamsForBusiness",
    }

    if attendees:
        attendee_list = []
        for email in attendees:
            # Create an attendee object
            # https://docs.microsoft.com/graph/api/resources/attendee?view=graph-rest-1.0
            attendee_list.append({
                'type': 'required',
                'emailAddress': { 'address': email }
            })

        new_event['attendees'] = attendee_list

    if body:
        # Create an itemBody object
        # https://docs.microsoft.com/graph/api/resources/itembody?view=graph-rest-1.0
        new_event['body'] = {
            'contentType': 'text',
            'content': body
        }

    # Set headers
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    response = requests.post(f'{GRAPH_URL}/me/events',
        headers=headers,
        data=json.dumps(new_event))
    if response.status_code != 201:
        print("Failed to create event:", response.status_code, response.text)

    return response

#/* spell-checker: disable */
# Basic lookup for mapping Windows time zone identifiers to
# IANA identifiers
# Mappings taken from
# https://github.com/unicode-org/cldr/blob/master/common/supplemental/windowsZones.xml
zone_mappings = {
    'Dateline Standard Time': 'Etc/GMT+12',
    'UTC-11': 'Etc/GMT+11',
    'Aleutian Standard Time': 'America/Adak',
    'Hawaiian Standard Time': 'Pacific/Honolulu',
    'Marquesas Standard Time': 'Pacific/Marquesas',
    'Alaskan Standard Time': 'America/Anchorage',
    'UTC-09': 'Etc/GMT+9',
    'Pacific Standard Time (Mexico)': 'America/Tijuana',
    'UTC-08': 'Etc/GMT+8',
    'Pacific Standard Time': 'America/Los_Angeles',
    'US Mountain Standard Time': 'America/Phoenix',
    'Mountain Standard Time (Mexico)': 'America/Chihuahua',
    'Mountain Standard Time': 'America/Denver',
    'Central America Standard Time': 'America/Guatemala',
    'Central Standard Time': 'America/Chicago',
    'Easter Island Standard Time': 'Pacific/Easter',
    'Central Standard Time (Mexico)': 'America/Mexico_City',
    'Canada Central Standard Time': 'America/Regina',
    'SA Pacific Standard Time': 'America/Bogota',
    'Eastern Standard Time (Mexico)': 'America/Cancun',
    'Eastern Standard Time': 'America/New_York',
    'Haiti Standard Time': 'America/Port-au-Prince',
    'Cuba Standard Time': 'America/Havana',
    'US Eastern Standard Time': 'America/Indianapolis',
    'Turks And Caicos Standard Time': 'America/Grand_Turk',
    'Paraguay Standard Time': 'America/Asuncion',
    'Atlantic Standard Time': 'America/Halifax',
    'Venezuela Standard Time': 'America/Caracas',
    'Central Brazilian Standard Time': 'America/Cuiaba',
    'SA Western Standard Time': 'America/La_Paz',
    'Pacific SA Standard Time': 'America/Santiago',
    'Newfoundland Standard Time': 'America/St_Johns',
    'Tocantins Standard Time': 'America/Araguaina',
    'E. South America Standard Time': 'America/Sao_Paulo',
    'SA Eastern Standard Time': 'America/Cayenne',
    'Argentina Standard Time': 'America/Buenos_Aires',
    'Greenland Standard Time': 'America/Godthab',
    'Montevideo Standard Time': 'America/Montevideo',
    'Magallanes Standard Time': 'America/Punta_Arenas',
    'Saint Pierre Standard Time': 'America/Miquelon',
    'Bahia Standard Time': 'America/Bahia',
    'UTC-02': 'Etc/GMT+2',
    'Azores Standard Time': 'Atlantic/Azores',
    'Cape Verde Standard Time': 'Atlantic/Cape_Verde',
    'UTC': 'Etc/GMT',
    'GMT Standard Time': 'Europe/London',
    'Greenwich Standard Time': 'Atlantic/Reykjavik',
    'Sao Tome Standard Time': 'Africa/Sao_Tome',
    'Morocco Standard Time': 'Africa/Casablanca',
    'W. Europe Standard Time': 'Europe/Berlin',
    'Central Europe Standard Time': 'Europe/Budapest',
    'Romance Standard Time': 'Europe/Paris',
    'Central European Standard Time': 'Europe/Warsaw',
    'W. Central Africa Standard Time': 'Africa/Lagos',
    'Jordan Standard Time': 'Asia/Amman',
    'GTB Standard Time': 'Europe/Bucharest',
    'Middle East Standard Time': 'Asia/Beirut',
    'Egypt Standard Time': 'Africa/Cairo',
    'E. Europe Standard Time': 'Europe/Chisinau',
    'Syria Standard Time': 'Asia/Damascus',
    'West Bank Standard Time': 'Asia/Hebron',
    'South Africa Standard Time': 'Africa/Johannesburg',
    'FLE Standard Time': 'Europe/Kiev',
    'Israel Standard Time': 'Asia/Jerusalem',
    'Kaliningrad Standard Time': 'Europe/Kaliningrad',
    'Sudan Standard Time': 'Africa/Khartoum',
    'Libya Standard Time': 'Africa/Tripoli',
    'Namibia Standard Time': 'Africa/Windhoek',
    'Arabic Standard Time': 'Asia/Baghdad',
    'Turkey Standard Time': 'Europe/Istanbul',
    'Arab Standard Time': 'Asia/Riyadh',
    'Belarus Standard Time': 'Europe/Minsk',
    'Russian Standard Time': 'Europe/Moscow',
    'E. Africa Standard Time': 'Africa/Nairobi',
    'Iran Standard Time': 'Asia/Tehran',
    'Arabian Standard Time': 'Asia/Dubai',
    'Astrakhan Standard Time': 'Europe/Astrakhan',
    'Azerbaijan Standard Time': 'Asia/Baku',
    'Russia Time Zone 3': 'Europe/Samara',
    'Mauritius Standard Time': 'Indian/Mauritius',
    'Saratov Standard Time': 'Europe/Saratov',
    'Georgian Standard Time': 'Asia/Tbilisi',
    'Volgograd Standard Time': 'Europe/Volgograd',
    'Caucasus Standard Time': 'Asia/Yerevan',
    'Afghanistan Standard Time': 'Asia/Kabul',
    'West Asia Standard Time': 'Asia/Tashkent',
    'Ekaterinburg Standard Time': 'Asia/Yekaterinburg',
    'Pakistan Standard Time': 'Asia/Karachi',
    'Qyzylorda Standard Time': 'Asia/Qyzylorda',
    'India Standard Time': 'Asia/Calcutta',
    'Sri Lanka Standard Time': 'Asia/Colombo',
    'Nepal Standard Time': 'Asia/Katmandu',
    'Central Asia Standard Time': 'Asia/Almaty',
    'Bangladesh Standard Time': 'Asia/Dhaka',
    'Omsk Standard Time': 'Asia/Omsk',
    'Myanmar Standard Time': 'Asia/Rangoon',
    'SE Asia Standard Time': 'Asia/Bangkok',
    'Altai Standard Time': 'Asia/Barnaul',
    'W. Mongolia Standard Time': 'Asia/Hovd',
    'North Asia Standard Time': 'Asia/Krasnoyarsk',
    'N. Central Asia Standard Time': 'Asia/Novosibirsk',
    'Tomsk Standard Time': 'Asia/Tomsk',
    'China Standard Time': 'Asia/Shanghai',
    'North Asia East Standard Time': 'Asia/Irkutsk',
    'Singapore Standard Time': 'Asia/Singapore',
    'W. Australia Standard Time': 'Australia/Perth',
    'Taipei Standard Time': 'Asia/Taipei',
    'Ulaanbaatar Standard Time': 'Asia/Ulaanbaatar',
    'Aus Central W. Standard Time': 'Australia/Eucla',
    'Transbaikal Standard Time': 'Asia/Chita',
    'Tokyo Standard Time': 'Asia/Tokyo',
    'North Korea Standard Time': 'Asia/Pyongyang',
    'Korea Standard Time': 'Asia/Seoul',
    'Yakutsk Standard Time': 'Asia/Yakutsk',
    'Cen. Australia Standard Time': 'Australia/Adelaide',
    'AUS Central Standard Time': 'Australia/Darwin',
    'E. Australia Standard Time': 'Australia/Brisbane',
    'AUS Eastern Standard Time': 'Australia/Sydney',
    'West Pacific Standard Time': 'Pacific/Port_Moresby',
    'Tasmania Standard Time': 'Australia/Hobart',
    'Vladivostok Standard Time': 'Asia/Vladivostok',
    'Lord Howe Standard Time': 'Australia/Lord_Howe',
    'Bougainville Standard Time': 'Pacific/Bougainville',
    'Russia Time Zone 10': 'Asia/Srednekolymsk',
    'Magadan Standard Time': 'Asia/Magadan',
    'Norfolk Standard Time': 'Pacific/Norfolk',
    'Sakhalin Standard Time': 'Asia/Sakhalin',
    'Central Pacific Standard Time': 'Pacific/Guadalcanal',
    'Russia Time Zone 11': 'Asia/Kamchatka',
    'New Zealand Standard Time': 'Pacific/Auckland',
    'UTC+12': 'Etc/GMT-12',
    'Fiji Standard Time': 'Pacific/Fiji',
    'Chatham Islands Standard Time': 'Pacific/Chatham',
    'UTC+13': 'Etc/GMT-13',
    'Tonga Standard Time': 'Pacific/Tongatapu',
    'Samoa Standard Time': 'Pacific/Apia',
    'Line Islands Standard Time': 'Pacific/Kiritimati'
}

def get_iana_from_windows(windows_tz_name):
    if windows_tz_name in zone_mappings:
        return zone_mappings[windows_tz_name]

    # Assume if not found value is
    # already an IANA name
    return windows_tz_name
