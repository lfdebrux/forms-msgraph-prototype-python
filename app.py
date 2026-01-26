import requests
from flask import Flask, render_template_string, request

from identity.flask import Auth

app = Flask(__name__)

app.config["SESSION_TYPE"] = "filesystem" # identity.flask.Auth expects Flask-Session to be configured
app.config.from_prefixed_env(prefix="SETTINGS_") # load config from any envvar starting SETTINGS__

auth = Auth(
    app,
    authority=f"https://login.microsoftonline.com/{app.config['GRAPH_API']['TENANT_ID']}",
    client_id=app.config["GRAPH_API"]["CLIENT_ID"],
    client_credential=app.config["GRAPH_API"]["CLIENT_SECRET"],
    redirect_uri="http://localhost:5000/auth/callback",
)

excel_scopes = ["Files.ReadWrite.AppFolder"]

@app.route("/")
def index():
    return f"""
        <h1>GOV.UK Forms prototype - Sending submissions to an Excel spreadsheet</h1>

        <p>
            The following links require you to be signed in to a Microsoft account,
            and may ask your consent for various permissions.
            <br />
            Don’t agree to anything you’re not comfortable with!
        </p>

        <a href="/excel/new">Create a new Excel spreadsheet</a>
    """

@app.route("/excel/new")
@auth.login_required(scopes=excel_scopes)
def excel_new(*, context):
    response = requests.get(
        "https://graph.microsoft.com/v1.0/me/drive/special/approot",
        headers={"Authorization": f"Bearer {context['access_token']}"},
        timeout=30,
    )
    response.raise_for_status()

    approot = response.json()
    #print({"approot": approot})

    drive_id = approot["parentReference"]["driveId"]

    return render_template_string("""
        <h1>Create a new Excel spreadsheet</h1>

        <p>
            This will create a new Excel spreadsheet,
            in a special app folder in your personal OneDrive
        </p>

        <form action="/excel/create" method="post">
            <input type="hidden" name="drive" value="{{ drive_id }}" />

            <button>Create Excel spreadsheet</button>
        </form>
    """, drive_id=drive_id)

@app.post("/excel/create")
@auth.login_required(scopes=excel_scopes)
def excel_create(*, context):
    with requests.Session() as session:
        session.headers.update({"Authorization": f"Bearer {context['access_token']}"})

        drive_id = request.form['drive']

        if not drive_id:
            return redirect("/new")

        response = session.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/special/approot",
            timeout=30
        )
        response.raise_for_status()

        approot = response.json()
        #print({"approot": approot})

        approot_drive_item_id = approot["id"]
        form_name = test_form["name"]
        file_name = f"{ form_name }.xlsx"

        # This is documented in https://learn.microsoft.com/en-us/answers/questions/830336/is-there-any-ms-graph-api-to-create-workbook-in-gi#answer-1348868 (and nowhere else?)
        response = session.post(
            f"https://graph.microsoft.com/v1.0/drive/items/{approot_drive_item_id}/children",
            timeout=30,
            json={
                "name": file_name,
                "file": { },
                "@microsoft.graph.conflictBehavior": "replace", # CHANGEME
            }
        )
        response.raise_for_status()

        file = response.json()
        #print({ "file": file })

        file_drive_item_url = f"https://graph.microsoft.com/v1.0/drive/items/{file['id']}"

        breakpoint()

        # Before we do anything else, we should set the permissions of the new file so that this app can access it using the
        # Files.SelectedOperations.Selected scope. Doing this means that a) the app only needs the File.ReadWrite.AppFolder
        # scope as a delegated permission, b) the form processor can (in theory) move the file out of the app folder and we
        # can still access it, and c) administrators can revoke access to individual files as they wish.
        response = session.post(
            f"{file_drive_item_url}/permissions",
            timeout=30,
            json={
                "roles": ["write"],
                "grantedTo": {
                    "application": {
                        "id": app.config["GRAPH_API"]["CLIENT_ID"],
                    },
                },
            },
        )
        response.raise_for_status()

        # Working with Excel is better with a worbook session, see https://learn.microsoft.com/en-gb/graph/workbook-best-practice.
        # Oh, one other thing to note, for this and other calls to Excel APIs, the docs say that occassionally there can be a 504 error,
        # and the correct response is to retry the request. Don't believe me? See https://learn.microsoft.com/en-us/graph/api/workbook-createsession#error-handling
        # Anyway, for this prototype we're not going to worry about that, we'll just raise an exception in case of any HTTP errors,
        # but I thought it was worth noting it down.
        response = session.post(
            f"{file_drive_item_url}/workbook/createSession",
            timeout=30,
            json={ "persistChanges": True }
        )
        response.raise_for_status()

        session.headers.update({ "workbook-session-id": response.json()["id"] })

        try:
            # Prepare spreadsheet to have submission data sent to it
            form_question_texts = [page["question_text"] for page in test_form["pages"]]
            headers = ["Reference", "Submitted at", *form_question_texts]
            sheet_name = "Sheet1"

            # There's probably a smarter way to do this...?
            # this only works for forms with fewer than 24 questions.
            # I can't figure out how to use R1C1 notation however
            header_end_column = chr(ord("A") + len(headers) - 1)
            header_address = f"A1:{header_end_column}1"

            response = session.patch(
                f"{file_drive_item_url}/workbook/worksheets/{sheet_name}/range(address='{header_address}')",
                json={
                    "values": [headers],
                },
            )
            response.raise_for_status()

            response = session.post(
                f"{file_drive_item_url}/workbook/tables/add",
                timeout=30,
                json={
                    "address": f"{sheet_name}!{header_address}",
                    "hasHeaders": True,
                },
            )
            response.raise_for_status()

            table = response.json()
            #print({"table": table})
        finally:
            response = session.post(
                f"{file_drive_item_url}/workbook/closeSession",
                timeout=30,
                json={},
            )

        return f"""
            <h1>Created Excel spreadsheet</h1>

            <p>
                Table ID: {table['id']}
            </p>

            <a href="{file['webUrl']}">{file['name']}</a>
        """

# For now use a fake form record
test_form = {
    "id": 1,
    "name": "Test form",
    "pages": [
        { "question_text": "What’s your name?", },
        { "question_text": "When’s your date of birth?", },
        { "question_text": "What’s your address?", },
    ],
}
