import json
import os
from flask import Flask, render_template, session, request, redirect, url_for
from flask_session import Session  # https://pythonhosted.org/Flask-Session
import msal
import app_config
from Utils import SharePointUtil

app = Flask(__name__)
app.config.from_object(app_config)
Session(app)
sp_util = SharePointUtil()

# This section is needed for url_for("foo", _external=True) to automatically
# generate http scheme when this sample is running on localhost,
# and to generate https scheme when it is deployed behind reversed proxy.
# See also https://flask.palletsprojects.com/en/1.0.x/deploying/wsgi-standalone/#proxy-setups
from werkzeug.middleware.proxy_fix import ProxyFix

app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)


@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template('index.html', user=session["user"], version=msal.__version__)


@app.route("/login")
def login():
    # Technically we could use empty list [] as scopes to do just sign in,
    # here we choose to also collect end user consent upfront
    session["flow"] = sp_util.build_auth_code_flow(scopes=app_config.SCOPE)
    return render_template("login.html", auth_url=session["flow"]["auth_uri"], version=msal.__version__)


@app.route(app_config.REDIRECT_PATH)  # Its absolute URL must match your app's redirect_uri set in AAD
def authorized():
    try:
        cache = sp_util.load_cache()
        result = sp_util.build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args)
        if "error" in result:
            return render_template("auth_error.html", result=result)
        session["user"] = result.get("id_token_claims")
        sp_util.save_cache(cache)
    except ValueError:  # Usually caused by CSRF
        pass  # Simply ignore them
    return redirect(url_for("index"))


@app.route("/logout")
def logout():
    session.clear()  # Wipe out user and its token cache from session
    return redirect(  # Also logout from your tenant's web session
        app_config.AUTHORITY + "/oauth2/v2.0/logout" +
        "?post_logout_redirect_uri=" + url_for("index", _external=True))


@app.route("/graphcall")
def graphcall():
    token = sp_util.get_token_from_cache(app_config.SCOPE)
    sp_files_res = list()
    if not token:
        return redirect(url_for("login"))
    sp_util.request_headers = {'Authorization': 'Bearer ' + token['access_token']}
    sites_resp = sp_util.make_share_point_api_call(app_config.GRAPH_API_CALLS.get("get_all_sites"))
    sites = list()
    if sites_resp.status_code == 200:
        os.makedirs(app_config.SHAREPOINT_FILE_DIR, exist_ok=True)
        sites = sites_resp.json().get("value", [])
    else:
        print("unable to find sharepoint sites!!")
        print(sites_resp.text)
    for site_data in sites:
        print("Fetching Data from site: ", site_data.get("displayName"))
        site_files_data = sp_util.process_sharepoint_site(site_data)
        sp_files_res.extend(site_files_data)
        sub_sites = sp_util.get_subsites(site_data).get("value", [])
        for sub_site_data in sub_sites:
            print("Fetching Data from subsite: ", sub_site_data.get("displayName"))
            site_files_data = sp_util.process_sharepoint_site(sub_site_data)
            sp_files_res.extend(site_files_data)
    with open(app_config.SHAREPOINT_CONTENT_JSON_FILE, "w") as f:
        f.write(json.dumps(sp_files_res, indent=4))
    print("Total files extracted: ", len(sp_files_res))
    return render_template('display.html', result=sp_files_res)


app.jinja_env.globals.update(_build_auth_code_flow=sp_util.build_auth_code_flow)  # Used in template

if __name__ == "__main__":
    app.run()
