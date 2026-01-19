from flask import Flask

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
@auth.login_required
def authenticated(*, context):
    return f"""
        <h1>Authenticated</h1>

        <p>
            Name: {context["user"]["name"]} </br>
            Email: {context["user"]["preferred_username"]}
        </p>

        <form action="/create" method="post">
            <button>Create an Excel spreadsheet in the application folder</button>
        </form>
    """
