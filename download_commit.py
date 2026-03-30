#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import requests
from azure.identity import ClientSecretCredential
import dotenv
import yaml

dotenv.load_dotenv()

# =========================================
# 設定読み込み
# =========================================
with open(
    os.path.join(os.path.dirname(__file__), "config.yaml"), encoding="utf-8"
) as f:
    CONFIG = yaml.safe_load(f)

# =========================================
# SharePoint 設定（環境変数から取得）
# =========================================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

sp = CONFIG["sharepoint"]
SITE_HOST = sp["site_host"]
SITE_PATH = sp["site_path"]
GRAPH_BASE = sp["graph_base"]


# =========================================
# SharePoint authentication
# =========================================
def get_token():
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID, client_id=CLIENT_ID, client_secret=CLIENT_SECRET
    )
    token = credential.get_token("https://graph.microsoft.com/.default")
    return token.token


# =========================================
# Graph API: サイトID取得
# =========================================
def get_site_id(token):
    url = f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        print(f"✗ サイトID取得失敗: {resp.status_code}")
        print(resp.text[:500])
        raise SystemExit(1)
    site_id = resp.json()["id"]
    print(f"✓ サイトID: {site_id}")
    return site_id


# =========================================
# ダウンロード（Graph API パスベース）
# =========================================
def download_file(site_id, file_path, save_path, token):
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{file_path}:/content"
    headers = {"Authorization": f"Bearer {token}"}
    print(f"→ Downloading: {file_path}")

    resp = requests.get(url, headers=headers)

    if resp.status_code != 200:
        print(f"  ✗ Failed: {resp.status_code}")
        print(f"  {resp.text[:300]}")
        return False

    with open(save_path, "wb") as f:
        f.write(resp.content)

    print(f"  ✓ Saved to {save_path}")
    return True


# =========================================
# メイン処理
# =========================================
def main():
    token = get_token()
    site_id = get_site_id(token)

    for entry in CONFIG["files"]:
        file_path = entry["remote_path"]
        save_name = entry["save_name"]
        print(f"=== {save_name} {file_path}===")
        download_file(site_id, file_path, save_name, token)


if __name__ == "__main__":
    main()
