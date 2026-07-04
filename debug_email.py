import json
from email_reader import authenticate_gmail

service = authenticate_gmail()
query = "is:unread has:attachment filename:pdf"
results = service.users().messages().list(userId='me', q=query).execute()
messages = results.get('messages', [])

if messages:
    msg_id = messages[0]['id']
    message = service.users().messages().get(userId='me', id=msg_id).execute()
    payload = message['payload']
    
    def print_parts(part, indent=""):
        print(f"{indent}MIME: {part.get('mimeType')} | Filename: {part.get('filename')}")
        if 'parts' in part:
            for p in part['parts']:
                print_parts(p, indent + "  ")
        else:
            if part.get('mimeType') in ['text/plain', 'text/html']:
                data = part.get('body', {}).get('data', '')
                import base64
                if data:
                    try:
                        decoded = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                        print(f"{indent}Content: {repr(decoded)[:100]}")
                    except Exception as e:
                        print(f"{indent}Error decoding: {e}")

    print_parts(payload)
else:
    print("No unread emails found")
