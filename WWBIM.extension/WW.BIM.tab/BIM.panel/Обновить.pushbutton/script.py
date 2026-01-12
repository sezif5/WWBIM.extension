# -*- coding: utf-8 -*-
"""Обновить плагин WW.BIM из GitHub (без version.txt)"""

__title__ = "Обновить\nплагин"
__author__ = "WW.BIM"
__doc__ = "Скачивает последнюю версию плагина с GitHub"

import os
import sys
import shutil
import zipfile
from datetime import datetime

from pyrevit import script, forms

# Python 2/3 совместимость
if sys.version_info[0] >= 3:
    from urllib.request import urlretrieve, urlopen
    from urllib.error import URLError
else:
    from urllib import urlretrieve
    from urllib2 import urlopen, URLError

# ============ НАСТРОЙКИ ============
GITHUB_USER = "sezif5"
GITHUB_REPO = "WW.BIM"
BRANCH = "main"
EXTENSION_NAME = "WW.BIM.extension"
# ===================================


def get_extension_path():
    script_path = os.path.dirname(__file__)
    ext_path = script_path

    for _ in range(10):
        if ext_path.endswith(".extension"):
            return ext_path
        parent = os.path.dirname(ext_path)
        if parent == ext_path:
            break
        ext_path = parent

    return None


def get_local_last_update(ext_path):
    """Читаем дату последнего обновления из last_update.txt"""
    last_update_file = os.path.join(ext_path, "last_update.txt")
    if os.path.exists(last_update_file):
        with open(last_update_file, "r") as f:
            return f.read().strip()
    return None


def get_remote_commit_date():
    """Получаем дату последнего коммита из GitHub API"""
    url = "https://api.github.com/repos/{}/{}/commits/{}".format(
        GITHUB_USER, GITHUB_REPO, BRANCH
    )
    try:
        response = urlopen(url, timeout=10)
        import json

        data = json.loads(response.read().decode("utf-8"))
        commit_date = data["commit"]["committer"]["date"]
        return commit_date
    except:
        return None


def compare_dates(remote_date, local_date):
    """Сравниваем даты - возвращает True если remote новее"""
    if not local_date:
        return True

    try:
        remote = datetime.strptime(remote_date, "%Y-%m-%dT%H:%M:%SZ")
        local = datetime.strptime(local_date, "%Y-%m-%dT%H:%M:%SZ")
        return remote > local
    except:
        return True


def save_last_update(ext_path, date):
    """Сохраняем дату последнего обновления"""
    last_update_file = os.path.join(ext_path, "last_update.txt")
    with open(last_update_file, "w") as f:
        f.write(date)


def download_update(ext_path):
    """Скачиваем и устанавливаем обновление"""
    temp_dir = os.environ.get("TEMP", os.environ.get("TMP", "/tmp"))
    zip_path = os.path.join(temp_dir, "wwbim_update.zip")
    extract_path = os.path.join(temp_dir, "wwbim_update")

    for branch in [BRANCH, "master", "main"]:
        zip_url = "https://github.com/{}/{}/archive/refs/heads/{}.zip".format(
            GITHUB_USER, GITHUB_REPO, branch
        )
        try:
            urlretrieve(zip_url, zip_path)
            used_branch = branch
            break
        except:
            continue
    else:
        return False, "Не удалось скачать обновление. Проверьте интернет."

    try:
        if os.path.exists(extract_path):
            shutil.rmtree(extract_path)

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_path)

        extracted_folder = os.path.join(
            extract_path, "{}-{}".format(GITHUB_REPO, used_branch)
        )

        source_ext = None
        possible_ext = os.path.join(extracted_folder, EXTENSION_NAME)
        if os.path.exists(possible_ext):
            source_ext = possible_ext
        else:
            for item in os.listdir(extracted_folder):
                if item.endswith(".extension"):
                    source_ext = os.path.join(extracted_folder, item)
                    break

        if not source_ext:
            source_ext = extracted_folder

        skip_files = [
            "user_config.json",
            "local_settings.py",
            ".user",
            "last_update.txt",
        ]

        for item in os.listdir(source_ext):
            if item in skip_files:
                continue
            if item.startswith(".git"):
                continue

            src = os.path.join(source_ext, item)
            dst = os.path.join(ext_path, item)

            try:
                if os.path.isdir(src):
                    if os.path.exists(dst):
                        shutil.rmtree(dst)
                    shutil.copytree(src, dst)
                else:
                    shutil.copy2(src, dst)
            except Exception as e:
                print("Не удалось скопировать {}: {}".format(item, e))

        try:
            os.remove(zip_path)
            shutil.rmtree(extract_path)
        except:
            pass

        return True, "Обновление успешно установлено!"

    except Exception as e:
        return False, "Ошибка при установке: {}".format(str(e))


def force_update(ext_path):
    with forms.ProgressBar(title="Скачивание обновления...") as pb:
        pb.update_progress(20, 100)
        success, message = download_update(ext_path)
        pb.update_progress(100, 100)
    return success, message


# ============ ГЛАВНЫЙ КОД ============
if __name__ == "__main__":
    ext_path = get_extension_path()

    if not ext_path:
        ext_path = forms.pick_folder(title="Выберите папку плагина (.extension)")
        if not ext_path:
            forms.alert("Путь к плагину не выбран", exitscript=True)

    local_date = get_local_last_update(ext_path)
    remote_date = get_remote_commit_date()

    if not remote_date:
        if forms.alert(
            "Не удалось проверить дату коммита на сервере.\n\n"
            "Последнее обновление: {}\n\n"
            "Выполнить принудительное обновление?".format(local_date or "неизвестно"),
            yes=True,
            no=True,
        ):
            success, message = force_update(ext_path)
            if success:
                remote_date = get_remote_commit_date()
                if remote_date:
                    save_last_update(ext_path, remote_date)
                forms.alert(
                    "{}\n\nПерезапустите Revit для применения изменений.".format(
                        message
                    )
                )
            else:
                forms.alert(message, warn_icon=True)
        sys.exit()

    if not compare_dates(remote_date, local_date):
        result = forms.alert(
            "У вас актуальная версия\n\n"
            "Последнее обновление: {}\n\n"
            "Выполнить принудительное обновление?".format(local_date or "неизвестно"),
            yes=True,
            no=True,
        )
        if not result:
            sys.exit()
    else:
        result = forms.alert(
            "Доступно обновление!\n\n"
            "Дата последнего коммита: {}\n\n"
            "Установить обновление?".format(remote_date),
            yes=True,
            no=True,
        )
        if not result:
            sys.exit()

    success, message = force_update(ext_path)

    if success:
        save_last_update(ext_path, remote_date)
        forms.alert(
            "{}\n\nПерезапустите Revit для применения изменений.".format(message)
        )
    else:
        forms.alert(message, warn_icon=True)
