import requests
from azure.identity import ClientSecretCredential
import os
import urllib.parse

S_EXCEL_PATH = "140.スケジュール・時間割/2025(R7)年度 時間割/【2025・04～09月 前期】全学年時間割.xlsx"
F_EXCEL_PATH = "140.スケジュール・時間割/2025(R7)年度 時間割/【2025・10～03月 後期】全学年時間割.xlsx"



CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")

HOSTNAME = os.getenv("HOSTNAME", "ncnj.sharepoint.com")
SITE_PATH = os.getenv("SITE_PATH", "/sites/staff_sharedfolders")

def get_access_token():
    """アクセストークンを取得"""
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    token = credential.get_token("https://graph.microsoft.com/.default")
    return token.token


def make_graph_api_request(token, base_url, path_params=None, operation_name="API呼び出し", include_accept_header=True):
    """Microsoft Graph APIへのGETリクエストを実行し、レスポンスを処理する共通関数
    
    Args:
        token (str): アクセストークン
        base_url (str): ベースURL
        path_params (str, optional): パスパラメータ（URLエンコードが必要な場合）
        operation_name (str): 操作名（ログ出力用）
        include_accept_header (bool): Accept: application/jsonヘッダーを含めるかどうか
    
    Returns:
        requests.Response or None: 成功時はResponseオブジェクト、失敗時はNone
    """
    # URLを構築
    if path_params:
        encoded_path = urllib.parse.quote(path_params)
        url = f"{base_url}{encoded_path}"
    else:
        url = base_url
    
    # ヘッダーを構築
    headers = {'Authorization': f'Bearer {token}'}
    if include_accept_header:
        headers['Accept'] = 'application/json'
    
    try:
        response = requests.get(url, headers=headers)
        print(f"api:{operation_name}結果: {response.status_code}")
        
        if response.status_code == 200:
            return response
        else:
            print(f"api:✗ {operation_name}失敗: {response.text[:200]}")
            return None
            
    except Exception as e:
        print(f"✗ {operation_name}エラー: {e}")
        return None


def get_site_id_from_url(token, hostname, site_path):
    """サイトパスからサイトIDを取得"""
    base_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    response = make_graph_api_request(token, base_url, operation_name="サイト情報取得")
    if response:
        site_data = response.json()
        site_id = site_data['id']
        print(f"✓ サイトID取得成功: {site_id}")
        return site_id
    else:
        return None

def get_file_by_path(token, site_id, file_path):
    """ファイルパスからファイル情報を取得"""
    base_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
    response = make_graph_api_request(token, base_url, path_params=file_path, operation_name="ファイル情報取得")
    if response:
        file_data = response.json()
        file_id = file_data['id']
        print(f"✓ ファイル情報取得成功: {file_id}")
        return file_data
    else:
        return None

def download_file_by_id(token, site_id, file_id, output_filename):
    """ファイルIDでファイルをダウンロード"""
    base_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}/content"
    response = make_graph_api_request(token, base_url, operation_name="ダウンロード", include_accept_header=False)
    if response:
        content = response.content
        with open(output_filename, 'wb') as f:
            f.write(content)
        print(f"✓ ダウンロード成功: {len(content)} bytes")
        return True
    else:
        return None

def download_excel_from_sharepoint_url(hostname, site_path, file_path, output_filename=None):
    """SharePointのURLから直接Excelファイルをダウンロード"""

    # 1. アクセストークンを取得
    token = get_access_token()
    # 2. サイトIDを取得
    site_id = get_site_id_from_url(token, hostname, site_path)
    
    if not site_id:
        print("✗ サイトIDの取得に失敗しました")
        return None
    
    # 3. ファイル情報を取得
    file_data = get_file_by_path(token, site_id, file_path)
    if not file_data:
        print("✗ ファイル情報の取得に失敗しました")
        return None
    
    # 4. ファイルをダウンロード
    download_file_by_id(token, site_id, file_data['id'], output_filename)


if __name__ == "__main__":
    # 前期 R7.09.04 Add by SUGIE
    download_excel_from_sharepoint_url(
        HOSTNAME, SITE_PATH, S_EXCEL_PATH , "schedule_spring.xlsx"
    )

    # 後期 R7.09.04 Add by SUGIE
    download_excel_from_sharepoint_url(
        HOSTNAME, SITE_PATH, F_EXCEL_PATH , "schedule_fall.xlsx"
    )
