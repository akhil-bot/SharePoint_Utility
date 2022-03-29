import requests
from flask import session, redirect, url_for
import os
import msal
from tika import parser
import traceback

import app_config


class SharePointUtil(object):
    def __init__(self, request_headers=None):
        self.request_headers = request_headers
        pass

    def make_share_point_api_call(self, sp_api):
        response = requests.get(sp_api, headers=self.request_headers, stream=True)
        status_code = response.status_code
        if status_code in [401, 403]:
            print("Refreshing Token since the current token has expired!!")
            token = self.get_token_from_cache(app_config.SCOPE)
            if not token:
                return redirect(url_for("login"))
            self.request_headers = {'Authorization': 'Bearer ' + token['access_token']}
            response = requests.get(sp_api, headers=self.request_headers)
        return response

    def get_file_details(self, sp_file_resp):
        file_data = dict()
        file_data["sp_file_id"] = sp_file_resp.get("id", "")
        sp_remote_item = sp_file_resp.get("remoteItem", {})
        sp_shared_details = sp_remote_item.get("shared", {})
        file_data["sp_file_download_url"] = sp_file_resp.get("@microsoft.graph.downloadUrl",
                                                             sp_remote_item.get("webDavUrl"))
        file_data["sp_file_web_url"] = sp_file_resp.get("webUrl", "")
        file_data["sp_file_name"] = sp_file_resp.get("name", "")
        file_data["sp_file_mime_type"] = sp_file_resp.get("file", {}).get("mimeType", "")
        if sp_remote_item:
            file_data["sp_drive_id"] = sp_remote_item.get("parentReference", {}).get("driveId", "")
            file_data["sp_drive_type"] = sp_remote_item.get("parentReference", {}).get("driveType", "")
        else:
            file_data["sp_drive_id"] = sp_file_resp.get("parentReference", {}).get("driveId")
            file_data["sp_drive_type"] = sp_file_resp.get("parentReference", {}).get("driveType", "")

        file_data["sp_file_created_date"] = sp_file_resp.get("createdDateTime", "")
        file_data["sp_file_lmod_date"] = sp_file_resp.get("lastModifiedDateTime", "")
        file_data["sp_file_created_by"] = sp_file_resp.get("createdBy", {}).get("user", {}).get("displayName", "")
        file_data["sp_file_lmod_by"] = sp_file_resp.get("lastModifiedBy", {}).get("user", {}).get("displayName", "")
        file_data["sp_file_size"] = sp_file_resp.get("size", 0)
        file_data["sp_content_type"] = "file"
        if sp_shared_details:
            file_data["sp_file_shared_by"] = sp_shared_details.get("sharedBy", {}).get("user", {}) \
                .get("displayName", "")
            file_data["sp_file_shared_date"] = sp_shared_details.get("sharedDateTime")
        return file_data

    def download_file(self, sp_file_data):
        drive_id = sp_file_data.get("sp_drive_id")
        file_id = sp_file_data.get("sp_file_id")
        file_download_url = f'{app_config.GRAPH_ENDPOINT}/drives/{drive_id}/items/{file_id}/content'
        file_name = sp_file_data.get("sp_file_name", "")
        print("Downloading File: ", file_name)
        response = self.make_share_point_api_call(file_download_url)
        if response.status_code == 200:
            local_file_path = os.path.join(app_config.SHAREPOINT_FILE_DIR, file_name)
            with open(local_file_path, "wb") as local_file:
                for chunk in response.iter_content(chunk_size=1024):
                    # writing one chunk at a time to file
                    if chunk:
                        local_file.write(chunk)
            return local_file_path, 200
        else:
            print("Unable to download the file: ", file_name)
            print(response.text)
            return "", 400

    def extract_file_content(self, sp_file_data):
        local_file_path, status_code = self.download_file(sp_file_data)
        if status_code == 200:
            tika_service_endpoint = app_config.TIKA_SERVICE_URL
            try:
                print("Extracting file: ", local_file_path)
                parsed_data = parser.from_file(local_file_path, tika_service_endpoint, xmlContent=False)
                # shutil.rmtree(app_config.SHAREPOINT_FILE_DIR)
                content = parsed_data.get("content", None)
                os.remove(local_file_path)
                if content:
                    content = content.strip()
                return content
            except:
                os.remove(local_file_path)
                print(traceback.format_exc())
                return None

        return None

    def extract_files_from_folder(self, sp_resp):
        drive_id = sp_resp.get("parentReference", {}).get("driveId", "")
        folder_id = sp_resp.get("id", "")
        get_files_url = f'{app_config.GRAPH_ENDPOINT}/drives/{drive_id}/items/{folder_id}/children'
        get_files_resp = self.make_share_point_api_call(get_files_url).json()
        folder_files_content = list()
        for sp_file_resp in get_files_resp.get("value", []):
            file_data = self.get_file_details(sp_file_resp)
            mime_type = file_data.get("sp_file_mime_type")
            if mime_type in ["application/pdf",
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document"]:
                file_content = self.extract_file_content(file_data)
                file_data["sp_file_content"] = file_content
                folder_files_content.append(file_data)
        return folder_files_content

    def load_cache(self):
        cache = msal.SerializableTokenCache()
        if session.get("token_cache"):
            cache.deserialize(session["token_cache"])
        return cache

    def save_cache(self, cache):
        if cache.has_state_changed:
            session["token_cache"] = cache.serialize()

    def build_msal_app(self, cache=None, authority=None):
        if not authority:
            authority = app_config.AUTHORITY
        return msal.ConfidentialClientApplication(
            app_config.CLIENT_ID, authority=authority,
            client_credential=app_config.CLIENT_SECRET, token_cache=cache)

    def build_auth_code_flow(self, authority=None, scopes=None):
        return self.build_msal_app(authority=authority).initiate_auth_code_flow(
            scopes or [],
            redirect_uri=url_for("authorized", _external=True))

    def get_token_from_cache(self, scope=None):
        cache = self.load_cache()  # This web app maintains one cache per session
        cca = self.build_msal_app(cache=cache)
        accounts = cca.get_accounts()
        if accounts:  # So all account(s) belong to the current signed-in user
            result = cca.acquire_token_silent(scope, account=accounts[0])
            self.save_cache(cache)
            return result

    def get_site_lists(self, site_id):
        site_lists_api = app_config.GRAPH_API_CALLS.get("get_lists")
        site_lists_api = site_lists_api.replace("{site_id}", site_id)
        site_lists_resp = self.make_share_point_api_call(site_lists_api)
        if site_lists_resp.status_code == 200:
            return site_lists_resp.json()
        else:
            print(site_lists_resp.text)
            return dict()

    def get_site_list_items(self, site_id, list_id):
        site_list_items_api = app_config.GRAPH_API_CALLS.get("get_list_items")
        site_list_items_api = site_list_items_api.replace("{site_id}", site_id).replace("{list_id}", list_id)
        site_list_items_resp = self.make_share_point_api_call(site_list_items_api)
        if site_list_items_resp.status_code == 200:
            return site_list_items_resp.json()
        else:
            print(site_list_items_resp.text)
            return dict()

    def get_list_drive_item(self, site_id, list_id, item_id):
        site_drive_item_api = app_config.GRAPH_API_CALLS.get("get_drive_item")
        site_drive_item_api = site_drive_item_api.replace("{site_id}", site_id).replace("{list_id}", list_id) \
            .replace("{item_id}", item_id)
        site_drive_item_resp = self.make_share_point_api_call(site_drive_item_api)
        if site_drive_item_resp.status_code == 200:
            return site_drive_item_resp.json()
        else:
            print(site_drive_item_resp.text)
            return None

    def process_sharepoint_site(self, site_data):
        site_id = site_data.get("id")
        site_lists = self.get_site_lists(site_id)
        sp_site_results = list()
        for list_data in site_lists.get("value", []):
            if list_data.get("list", {}).get("template") in ['documentLibrary']:
                list_id = list_data.get("id")
                list_items = self.get_site_list_items(site_id, list_id)
                for list_item_data in list_items.get("value", []):
                    if list_item_data.get("contentType", {}).get("id", "").startswith("0x012000"):
                        continue
                    item_id = list_item_data.get("id")
                    list_drive_item = self.get_list_drive_item(site_id, list_id, item_id)
                    if list_drive_item:
                        sp_file_data = self.get_file_details(list_drive_item)
                        mime_type = sp_file_data.get("sp_file_mime_type", "")
                        if mime_type in ["application/pdf",
                                         "application/vnd.openxmlformats-officedocument.wordprocessingml.document"]:
                            file_content = self.extract_file_content(sp_file_data)
                            sp_file_data["sp_file_content"] = file_content
                            sp_file_data["sp_site_id"] = site_id
                            sp_file_data["sp_site_name"] = site_data.get("displayName", "")
                            sp_file_data["sp_site_url"] = site_data.get("webUrl")
                            sp_site_results.append(sp_file_data)
        return sp_site_results

    def get_subsites(self, site_data):
        site_id = site_data.get("id")
        get_subsites_api = app_config.GRAPH_API_CALLS.get("get_subsites")
        get_subsites_api = get_subsites_api.replace("{site_id}", site_id)
        subsites_resp = self.make_share_point_api_call(get_subsites_api)
        if subsites_resp.status_code == 200:
            return subsites_resp.json()
        else:
            print(subsites_resp.text)
            return dict()
