import fire
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/presentations"]  # read-write

# The ID of a sample presentation.
PRESENTATION_ID = "1jvtFxrzqQKiSCqeAm5Nggg-JAReXV5Hf4hjos1yNy9A"


def get_presentation(presentation_id):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    service = build("slides", "v1", credentials=creds)

    presentation = service.presentations().get(presentationId=presentation_id).execute()
    return service, presentation


def get_slide_notes_text(slide):
    slide_properties = slide.get("slideProperties")
    notes_page = slide_properties.get("notesPage", {})
    slide_notes_id = notes_page.get("notesProperties", {}).get("speakerNotesObjectId")
    page_elements = notes_page.get("pageElements")
    try:
        pe = next(pe for pe in page_elements if pe.get("objectId") == slide_notes_id)
    except StopIteration:
        return ""
    text_elements = pe.get("shape", {}).get("text", {}).get("textElements", [])
    text = "".join([te.get("textRun", {}).get("content", "") for te in text_elements])
    return text


def main(
    presentation_id=PRESENTATION_ID, remove_all=False, last_is_zero=False, color="LIGHT2",
    fmt="{n}",  # set the number format, e.g. "#{n}"
):
    """
    color is one of https://developers.google.com/slides/reference/rest/v1/presentations.pages/other#Page.ThemeColorType
    """
    slides_service, presentation = get_presentation(presentation_id)
    slides = presentation.get("slides")

    print(f"The presentation contains {len(slides)} slides")
    counted_slides = []
    delete_requests = []
    no_number = set()
    in_supplement = False
    for slide in slides:
        notes = get_slide_notes_text(slide)
        if "@supplement-start@" in notes:
            in_supplement = True
        if not "@skip-count@" in notes and not in_supplement:
            counted_slides.append(slide)
        if "@no-number@" in notes:
            no_number.add(slide.get("objectId"))
        remove_els = [
            el["objectId"] for el in slide.get("pageElements", [])
            if el["objectId"].startswith("sedils_elgoog")
        ]
        for el in remove_els:
            delete_requests.append({"deleteObject": {"objectId": el}})

    if delete_requests:
        response = slides_service.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": delete_requests},
        ).execute()
        # print(response.get("replies"))

    if remove_all:
        return

    total = len(counted_slides)
    print(f"Counting {total} slides")
    print(f"Not showing number on {len(no_number)} slides")
    requests = []
    for i, slide in enumerate(counted_slides):
        if slide.get("objectId") in no_number:
            continue
        n = total - i - last_is_zero
        element_id = f"sedils_elgoog_{n}"
        width = {"magnitude": 60, "unit": "PT"}
        height = {"magnitude": 35, "unit": "PT"}
        page_id = slide.get("objectId")
        requests.append({
            'createShape': {
                'objectId': element_id,
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': page_id,
                    'size': {'height': height, 'width': width},
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 655,
                        'translateY': 370,
                        'unit': 'PT'
                    },
                },
            },
        })
        requests.append(
            {'insertText': {'objectId': element_id, 'insertionIndex': 0, 'text': fmt.format(n=n)}}
        )
        requests.append({  # right-align
            "updateParagraphStyle": {
                "objectId": element_id, "fields": "alignment", "style": {"alignment": "END"},
            }
        })
        requests.append({  # set color
            "updateTextStyle": {
                "objectId": element_id,
                "fields": "foregroundColor",
                "style": {"foregroundColor": {"opaqueColor": {"themeColor": color}}},
            }
        })

    if requests:
        response = slides_service.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests},
        ).execute()
        # print(response.get("replies"))


if __name__ == "__main__":
    fire.Fire()
