#!/usr/bin/env python

import json
import uuid

import urllib.request as urllib_request
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode


SERVER_URL = "http://192.168.88.178/RevitServerAdmin2023"
TIMEOUT_SEC = 10


class RevitServerApi(object):
    def __init__(self, base_url, timeout_sec=10, user_name=None, machine_name=None):
        self.base_url = base_url.rstrip("/")
        self.api_base = self.base_url + "/api"
        self.timeout_sec = timeout_sec
        self.user_name = user_name
        self.machine_name = machine_name
        self.host = self._extract_host(base_url)

    def _extract_host(self, url):
        no_scheme = url.replace("http://", "").replace("https://", "")
        return no_scheme.split("/")[0]

    def _default_headers(self):
        headers = {"Accept": "application/json"}
        if self.user_name:
            headers["User-Name"] = self.user_name
        if self.machine_name:
            headers["User-Machine-Name"] = self.machine_name
        headers["Operation-GUID"] = str(uuid.uuid4())
        return headers

    def _request_json(self, path, query=None):
        url = self.api_base + path
        if query:
            url = url + "?" + urlencode(query)

        request = urllib_request.Request(url, headers=self._default_headers())
        response = urllib_request.urlopen(request, timeout=self.timeout_sec)
        try:
            status = response.getcode() if hasattr(response, "getcode") else None
            body = response.read().decode("utf-8", "ignore")
            payload = json.loads(body) if body else None
            return status, payload
        finally:
            response.close()

    def ping(self):
        request = urllib_request.Request(self.base_url)
        try:
            response = urllib_request.urlopen(request, timeout=self.timeout_sec)
            try:
                status = response.getcode() if hasattr(response, "getcode") else None
                return {"ok": True, "http_status": status}
            finally:
                response.close()
        except HTTPError as err:
            return {"ok": True, "http_status": err.code}
        except URLError as err:
            return {"ok": False, "error": str(err)}
        except Exception as err:
            return {"ok": False, "error": str(err)}

    def get_servers(self):
        _, payload = self._request_json(
            "/server/servers", {"id": self.host, "refresh": "true"}
        )
        return payload or []

    def get_root_server_id(self):
        servers = self.get_servers()
        if not servers:
            return None
        return servers[0].get("Id")

    def get_subitems(self, item_id, depth=2):
        _, payload = self._request_json(
            "/folder/SubItems", {"id": item_id, "depth": depth}
        )
        return payload

    def get_model_details(self, model_id):
        _, payload = self._request_json("/model/details", {"id": model_id})
        return payload

    def get_model_history(self, model_id):
        _, payload = self._request_json(
            "/model/ModelHistories", {"type": "rs-model", "id": model_id}
        )
        return payload

    def find_model(self, model_name, depth=10):
        server_id = self.get_root_server_id()
        if not server_id:
            return None

        tree = self.get_subitems(server_id, depth=depth)
        matches = []
        self._collect_models(tree, model_name, matches)
        if not matches:
            return None
        return matches[0]

    def list_models(self, depth=2):
        server_id = self.get_root_server_id()
        if not server_id:
            return []

        tree = self.get_subitems(server_id, depth=depth)
        models = []
        self._collect_all_models(tree, models)
        return models

    def _collect_models(self, node, model_name, matches):
        if not isinstance(node, dict):
            return

        item_type = node.get("Type")
        name = node.get("Name", "")
        if item_type == "rs-model" and name.lower() == model_name.lower():
            matches.append(node)

        for child in node.get("Children") or []:
            self._collect_models(child, model_name, matches)

    def _collect_all_models(self, node, models):
        if not isinstance(node, dict):
            return

        if node.get("Type") == "rs-model":
            models.append(node)

        for child in node.get("Children") or []:
            self._collect_all_models(child, models)

    def get_model_data(self, model_name, fields, include_history=False, depth=10):
        model_item = self.find_model(model_name=model_name, depth=depth)
        if not model_item:
            return {
                "ok": False,
                "error": "Model not found",
                "model_name": model_name,
            }

        model_id = model_item.get("Id")
        details = self.get_model_details(model_id)
        if not isinstance(details, dict):
            details = {}

        result = {
            "ok": True,
            "model_name": model_name,
            "model_id": model_id,
            "requested": {},
        }

        for field_name in fields or []:
            result["requested"][field_name] = details.get(field_name)

        if include_history:
            result["history"] = self.get_model_history(model_id)

        result["model"] = {
            "Name": model_item.get("Name"),
            "ServerPath": model_item.get("ServerPath"),
            "Type": model_item.get("Type"),
            "LockStatus": model_item.get("LockStatus"),
        }
        return result


def query_model(
    base_url, model_name, fields, include_history=False, timeout_sec=10, depth=10
):
    client = RevitServerApi(base_url=base_url, timeout_sec=timeout_sec)
    ping_result = client.ping()
    if not ping_result.get("ok"):
        return {
            "ok": False,
            "error": "Server is unreachable",
            "ping": ping_result,
        }

    try:
        data = client.get_model_data(
            model_name=model_name,
            fields=fields,
            include_history=include_history,
            depth=depth,
        )
        data["ping"] = ping_result
        return data
    except HTTPError as err:
        return {
            "ok": False,
            "error": "HTTP error",
            "http_status": err.code,
            "url": getattr(err, "url", None),
        }
    except URLError as err:
        return {
            "ok": False,
            "error": "URL error",
            "details": str(err),
        }
    except Exception as err:
        return {
            "ok": False,
            "error": "Unexpected error",
            "details": str(err),
        }


if __name__ == "__main__":
    demo_client = RevitServerApi(base_url=SERVER_URL, timeout_sec=TIMEOUT_SEC)
    models = demo_client.list_models(depth=4)

    if not models:
        print(
            json.dumps(
                {"ok": False, "error": "No models found"}, ensure_ascii=False, indent=2
            )
        )
    else:
        sample_model_name = models[0].get("Name")
        example_result = query_model(
            base_url=SERVER_URL,
            model_name=sample_model_name,
            fields=[
                "ModelSize",
                "SupportSize",
                "DateCreated",
                "DateModified",
                "LockStatus",
                "LockContext",
            ],
            include_history=False,
            timeout_sec=TIMEOUT_SEC,
            depth=6,
        )
        print(json.dumps(example_result, ensure_ascii=False, indent=2))
