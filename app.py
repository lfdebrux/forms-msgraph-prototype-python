import requests
from flask import Flask, render_template_string, request

from identity.flask import Auth

app = Flask(__name__)

app.config["SESSION_TYPE"] = "filesystem" # identity.flask.Auth expects Flask-Session to be configured
app.config.from_prefixed_env(prefix="SETTINGS_") # load config from any envvar starting SETTINGS__

auth = Auth(
    app,
    authority=f"https://login.microsoftonline.com/common",
    client_id=app.config["GRAPH_API"]["CLIENT_ID"],
    client_credential=app.config["GRAPH_API"]["CLIENT_SECRET"],
    redirect_uri="http://localhost:5000/auth/callback",
)

# Although we don't need all scopes for all requests, it's better to request everything up-front,
# otherwise there can be more than one consent page for the form creator
scopes = ["Files.ReadWrite.AppFolder"]

@app.route("/")
def index():
    return f"""
        <h1>Testing writing to Excel file</h1>

        <p>
            Tenant: {auth._authority} </br>
            Client ID: {auth._client_id}
        </p>

        <a href="/auth">Authenticate with Microsoft Graph</a>
    """

@app.route("/auth")
@auth.login_required(scopes=scopes)
def authenticated(*, context):
    return f"""
        <h1>Authenticated</h1>

        <p>
            Name: {context["user"]["name"]} </br>
            Email: {context["user"]["preferred_username"]}
        </p>

        <a href="/new">Create an Excel spreadsheet</a>
    """

@app.route("/new")
@auth.login_required(scopes=scopes)
def new(*, context):
    response = requests.get(
        "https://graph.microsoft.com/v1.0/me/drives",
        headers={"Authorization": f"Bearer {context['access_token']}"},
        timeout=30,
    )
    response.raise_for_status()

    drives = response.json()
    print({"drives": drives})

    # NOTE: we probably want to filter out OneDrive and ODCMetadataArchive
    # NOTE: might be better to look at sites rather than drives?
    drive_ids = {
        drive['name']: drive['id']
        for drive in drives['value']
    }

    return render_template_string("""
        <h1>Choose a drive to store the spreadsheet in</h1>

        <form action="/create" method="post">
            <fieldset>
                {% for drive_name, drive_id in drive_ids.items() %}
                <div>
                    <input type="radio" id="{{ drive_id }}" name="drive" value="{{ drive_id }}" />
                    <label for="{{ drive_id }}">{{ drive_name }}</label>
                </div>
                {% endfor %}
            </fieldset>

            <button>Create Excel spreadsheet here</button>
        </form>
    """, drive_ids=drive_ids)

@app.post("/create")
@auth.login_required(scopes=scopes)
def create(*, context):
    drive_id = request.form['drive']

    if not drive_id:
        return redirect("/new")

    response = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/special/approot",
        headers={"Authorization": f"Bearer {context['access_token']}"},
        timeout=30
    )
    response.raise_for_status()

    approot = response.json()
    print({"approot": approot})

    approot_item_id = approot["id"]
    form_name = "Test form"
    file_name = f"{ form_name }.xlsx"

    # This is documented in https://learn.microsoft.com/en-us/answers/questions/830336/is-there-any-ms-graph-api-to-create-workbook-in-gi#answer-1348868 (and nowhere else?)
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{approot_item_id}/children",
        headers={"Authorization": f"Bearer {context['access_token']}"},
        timeout=30,
        json={
            "name": file_name,
            "file": { },
            "@microsoft.graph.conflictBehavior": "rename",
        }
    )
    response.raise_for_status()

    print(response.json())

    return f"""
        <h1>Created Excel spreadsheet</h1>

        <p>
            {repr(response.json())}
        </p>
    """
