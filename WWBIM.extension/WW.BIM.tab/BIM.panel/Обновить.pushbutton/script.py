# -*- coding: utf-8 -*-
"""Обновление WWBIM.extension из GitHub."""

from __future__ import print_function

import json
import os
import re
import shutil
import sys
import tempfile
import traceback
import zipfile
from datetime import datetime

from pyrevit import forms, script

try:
    from System.Net import WebClient, ServicePointManager, SecurityProtocolType
    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
except Exception:
    WebClient = None


__title__ = u"Обновить\nWWBIM"
__author__ = u"WWBIM"
__doc__ = u"Проверяет и устанавливает последнюю версию WWBIM.extension с GitHub."

GITHUB_USER = "sezif5"
GITHUB_REPO = "WWBIM.extension"
BRANCHES = ("main", "master")
EXTENSION_NAME = "WWBIM.extension"
STATE_FILE = ".wwbim-update-state.json"
MANIFEST_FILE = ".wwbim-manifest.json"

# Эти файлы принадлежат локальной установке и не заменяются архивом GitHub.
PRESERVED_FILES = set([
    "user_config.json",
    "local_settings.py",
    ".user",
    "last_update.txt",
    STATE_FILE,
    MANIFEST_FILE,
])
IGNORED_RUNTIME_NAMES = set([".DS_Store", "Thumbs.db"])


def http_client():
    if WebClient is None:
        return None
    client = WebClient()
    client.Headers.Add("User-Agent", "WWBIM-Updater/2.0")
    return client


def http_get_json(url):
    client = http_client()
    if client is None:
        return None
    return json.loads(client.DownloadString(url))


def http_download(url, path):
    client = http_client()
    if client is None:
        raise RuntimeError(u"WebClient недоступен в текущем окружении pyRevit.")
    client.DownloadFile(url, path)


def get_extension_path():
    current = os.path.abspath(os.path.dirname(__file__))
    for unused in range(12):
        if current.endswith(".extension"):
            return current
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    return None


def read_json(path, default):
    try:
        with open(path, "r") as stream:
            value = json.load(stream)
        return value
    except Exception:
        return default


def local_state(ext_path):
    state = read_json(os.path.join(ext_path, STATE_FILE), {})
    if isinstance(state, dict):
        return state
    return {}


def remote_commit():
    errors = []
    for branch in BRANCHES:
        url = "https://api.github.com/repos/{}/{}/commits/{}".format(
            GITHUB_USER, GITHUB_REPO, branch)
        try:
            data = http_get_json(url)
            commit = data["commit"]
            return {
                "sha": data["sha"],
                "date": commit["committer"]["date"],
                "message": commit["message"].splitlines()[0],
                "branch": branch,
            }
        except Exception as error:
            errors.append("{}: {}".format(branch, error))
    raise RuntimeError(u"Не удалось получить данные GitHub: {}".format("; ".join(errors)))


def needs_update(remote, local):
    local_sha = local.get("sha")
    if local_sha:
        return local_sha.lower() != remote["sha"].lower()
    local_date = local.get("date") or local.get("last_update")
    if not local_date:
        return True
    try:
        remote_dt = datetime.strptime(remote["date"], "%Y-%m-%dT%H:%M:%SZ")
        local_dt = datetime.strptime(local_date, "%Y-%m-%dT%H:%M:%SZ")
        return remote_dt > local_dt
    except Exception:
        return True


def safe_extract(archive_path, destination):
    root = os.path.realpath(destination)
    if not os.path.isdir(destination):
        os.makedirs(destination)
    with zipfile.ZipFile(archive_path, "r") as archive:
        for info in archive.infolist():
            name = info.filename.replace("\\", "/")
            if name.startswith("/") or re.match(r"^[A-Za-z]:", name):
                raise RuntimeError(u"Архив содержит абсолютный путь: {}".format(name))
            parts = [part for part in name.split("/") if part not in ("", ".")]
            if ".." in parts:
                raise RuntimeError(u"Архив содержит небезопасный путь: {}".format(name))
            target = os.path.realpath(os.path.join(destination, *parts))
            if target != root and not target.startswith(root + os.sep):
                raise RuntimeError(u"Архив выходит за пределы временной папки.")
            archive.extract(info, destination)


def find_source_extension(extracted_root, branch):
    expected = os.path.join(extracted_root, "{}-{}".format(GITHUB_REPO, branch), EXTENSION_NAME)
    if os.path.isdir(expected):
        return expected
    archive_root = os.path.join(extracted_root, "{}-{}".format(GITHUB_REPO, branch))
    if not os.path.isdir(archive_root):
        raise RuntimeError(u"В архиве GitHub не найден корневой каталог репозитория.")
    for item in os.listdir(archive_root):
        candidate = os.path.join(archive_root, item)
        if item.endswith(".extension") and os.path.isdir(candidate):
            return candidate
    raise RuntimeError(u"В архиве GitHub не найден каталог .extension.")


def relative_files(source_root):
    result = []
    for root, directories, files in os.walk(source_root):
        directories[:] = [d for d in directories if not d.startswith(".git")]
        directories[:] = [d for d in directories if d != "__pycache__"]
        for filename in files:
            relative = os.path.relpath(os.path.join(root, filename), source_root)
            relative = relative.replace(os.sep, "/")
            if (relative in PRESERVED_FILES or filename.startswith(".git") or
                    filename in IGNORED_RUNTIME_NAMES or filename.endswith(".pyc")):
                continue
            result.append(relative)
    return sorted(result)


def copy_update(source_root, ext_path, remote):
    new_files = relative_files(source_root)
    if not new_files:
        raise RuntimeError(u"В архиве нет файлов расширения для установки.")

    old_manifest_path = os.path.join(ext_path, MANIFEST_FILE)
    old_manifest = read_json(old_manifest_path, [])
    if not isinstance(old_manifest, list):
        old_manifest = []
    new_set = set(new_files)
    errors = []

    # Удаляются только файлы, которые были установлены предыдущей версией.
    # Пользовательские файлы вне manifest не затрагиваются.
    for relative in old_manifest:
        if relative in new_set or relative in PRESERVED_FILES:
            continue
        target = os.path.join(ext_path, relative.replace("/", os.sep))
        try:
            if os.path.isfile(target):
                os.remove(target)
        except Exception as error:
            errors.append(u"удаление {}: {}".format(relative, error))

    for relative in new_files:
        source = os.path.join(source_root, relative.replace("/", os.sep))
        target = os.path.join(ext_path, relative.replace("/", os.sep))
        try:
            parent = os.path.dirname(target)
            if not os.path.isdir(parent):
                os.makedirs(parent)
            shutil.copy2(source, target)
        except Exception as error:
            errors.append(u"копирование {}: {}".format(relative, error))

    if errors:
        raise RuntimeError(u"Обновление выполнено не полностью:\n{}".format("\n".join(errors)))

    state = {
        "sha": remote["sha"],
        "date": remote["date"],
        "message": remote["message"],
        "branch": remote["branch"],
    }
    with open(os.path.join(ext_path, STATE_FILE), "w") as stream:
        json.dump(state, stream, indent=2)
    with open(os.path.join(ext_path, MANIFEST_FILE), "w") as stream:
        json.dump(new_files, stream, indent=2)
    with open(os.path.join(ext_path, "last_update.txt"), "w") as stream:
        stream.write(remote["date"])


def download_and_install(ext_path, remote):
    temp_root = tempfile.mkdtemp(prefix="wwbim-update-")
    archive_path = os.path.join(temp_root, "repository.zip")
    extracted_path = os.path.join(temp_root, "extracted")
    try:
        url = "https://github.com/{}/{}/archive/refs/heads/{}.zip".format(
            GITHUB_USER, GITHUB_REPO, remote["branch"])
        http_download(url, archive_path)
        safe_extract(archive_path, extracted_path)
        source_root = find_source_extension(extracted_path, remote["branch"])
        copy_update(source_root, ext_path, remote)
        return True, u"Установлен коммит {}: {}".format(remote["sha"][:8], remote["message"])
    finally:
        try:
            shutil.rmtree(temp_root)
        except Exception:
            pass


def main():
    ext_path = get_extension_path()
    if not ext_path:
        ext_path = forms.pick_folder(title=u"Выберите папку WWBIM.extension")
    if not ext_path:
        script.exit()
    if WebClient is None:
        forms.alert(u"WebClient недоступен. Обновление отменено.", warn_icon=True)
        script.exit()

    try:
        remote = remote_commit()
        local = local_state(ext_path)
        if not needs_update(remote, local):
            if not forms.alert(
                    u"Установлена актуальная версия.\n\nКоммит: {}\n{}\n\nПроверить обновление принудительно?".format(
                        remote["sha"][:8], remote["message"]), yes=True, no=True):
                return
        else:
            if not forms.alert(
                    u"Доступно обновление.\n\nНовый коммит: {}\n{}\n\nУстановить?".format(
                        remote["sha"][:8], remote["message"]), yes=True, no=True):
                return
        success, message = download_and_install(ext_path, remote)
        if success:
            forms.alert(message + u"\n\nПерезагрузите Revit или обновите pyRevit.")
    except Exception as error:
        forms.alert(u"Обновление не выполнено:\n\n{}\n\n{}".format(
            error, traceback.format_exc()), warn_icon=True)


if __name__ == "__main__":
    main()
