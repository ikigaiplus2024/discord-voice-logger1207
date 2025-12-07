import discord
from discord.ext import commands
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import os
import json

# 環境変数から設定を取得
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
DISCORD_TOKEN = os.environ.get('DISCORD_TOKEN')
CREDENTIALS_JSON = os.environ.get('CREDENTIALS_JSON')

# Google Sheets の設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_NAME = 'ボイスログ'

# 日本時間(JST)のタイムゾーン設定
JST = timezone(timedelta(hours=9))

# Google Sheets 認証(環境変数から)
def get_google_sheets_client():
    if CREDENTIALS_JSON:
        # 環境変数からJSON文字列を読み込み
        creds_dict = json.loads(CREDENTIALS_JSON)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        # ローカルファイルから読み込み(開発時用)
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    return gspread.authorize(creds)

# スプレッドシートの初期化
def initialize_sheet():
    try:
        client = get_google_sheets_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        
        # シートが存在するか確認、なければ作成
        try:
            sheet = spreadsheet.worksheet(SHEET_NAME)
        except gspread.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(title=SHEET_NAME, rows=1000, cols=6)
        
        # ヘッダー行を設定
        headers = ['日付', '名前', 'ID', '部屋の名前', '入室時間', '退出時間']
        if sheet.row_values(1) != headers:
            sheet.update([headers], 'A1:F1')
        
        print("✅ スプレッドシート初期化完了")
        return sheet
    except Exception as e:
        print(f"❌ スプレッドシート初期化エラー: {e}")
        return None

# ログをスプレッドシートに追加(入室時のみ使用)
def log_to_sheet(date, name, user_id, channel_name, join_time, leave_time=""):
    try:
        client = get_google_sheets_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        sheet = spreadsheet.worksheet(SHEET_NAME)
        
        # 新しい行を追加
        row = [date, name, str(user_id), channel_name, join_time, leave_time]
        sheet.append_row(row, value_input_option='USER_ENTERED')
        print(f"📝 入室記録: {name} - {channel_name} ({join_time})")
        
    except Exception as e:
        print(f"❌ スプレッドシート書き込みエラー: {e}")

# 退出時間を既存の行に更新
def update_leave_time(user_id, channel_name, leave_time):
    try:
        client = get_google_sheets_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        sheet = spreadsheet.worksheet(SHEET_NAME)
        
        # 全データを取得
        all_values = sheet.get_all_values()
        
        # 最後の行から遡って、該当ユーザーの入室記録を探す
        for i in range(len(all_values) - 1, 0, -1):  # 最後の行から検索(0行目はヘッダー)
            row = all_values[i]
            # C列(index 2)がID、D列(index 3)が部屋名、F列(index 5)が退出時間
            if len(row) >= 6:
                # IDと部屋名が一致し、退出時間が空欄の行を探す
                if row[2] == str(user_id) and row[3] == channel_name and (len(row) < 6 or row[5] == ""):
                    # F列(6列目)に退出時間を更新
                    sheet.update_cell(i + 1, 6, leave_time)
                    print(f"📝 退出記録: {row[1]} - {channel_name} ({leave_time})")
                    return True
        
        print(f"⚠️ 入室記録が見つかりません: UserID={user_id}, Channel={channel_name}")
        return False
        
    except Exception as e:
        print(f"❌ 退出時間更新エラー: {e}")
        return False

# 入室時間を記録する辞書
user_join_times = {}

# Discord Bot の設定
intents = discord.Intents.default()
intents.voice_states = True
intents.guilds = True
intents.members = True

bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print(f'✅ {bot.user} としてログインしました')
    print('👀 ボイスチャンネルの監視を開始します...')
    print(f'🕐 現在の日本時間: {datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")}')
    
    # スプレッドシートの初期化
    initialize_sheet()

@bot.event
async def on_voice_state_update(member, before, after):
    """ボイスチャンネルの入退室を検知"""
    
    # 日本時間を取得
    now = datetime.now(JST)
    date = now.strftime('%Y-%m-%d')
    time_str = now.strftime('%H:%M:%S')
    
    # 入室検知
    if before.channel is None and after.channel is not None:
        # 入室時間を記録
        key = f"{member.id}_{after.channel.id}"
        user_join_times[key] = time_str
        
        print(f"🟢 入室: {member.display_name} → {after.channel.name} ({time_str})")
        
        # スプレッドシートに記録(退出時間は空欄)
        log_to_sheet(
            date=date,
            name=member.display_name,
            user_id=member.id,
            channel_name=after.channel.name,
            join_time=time_str,
            leave_time=""
        )
    
    # 退出検知
    elif before.channel is not None and after.channel is None:
        key = f"{member.id}_{before.channel.id}"
        
        print(f"🔴 退出: {member.display_name} ← {before.channel.name} ({time_str})")
        
        # 既存の行に退出時間を更新
        update_leave_time(
            user_id=member.id,
            channel_name=before.channel.name,
            leave_time=time_str
        )
        
        # 記録を削除
        if key in user_join_times:
            del user_join_times[key]
    
    # チャンネル移動検知
    elif before.channel is not None and after.channel is not None and before.channel != after.channel:
        # 前のチャンネルから退出
        key_before = f"{member.id}_{before.channel.id}"
        
        print(f"🔄 移動: {member.display_name} {before.channel.name} → {after.channel.name}")
        
        # 前のチャンネルの退出時間を更新
        update_leave_time(
            user_id=member.id,
            channel_name=before.channel.name,
            leave_time=time_str
        )
        
        if key_before in user_join_times:
            del user_join_times[key_before]
        
        # 新しいチャンネルへの入室を記録
        key_after = f"{member.id}_{after.channel.id}"
        user_join_times[key_after] = time_str
        
        log_to_sheet(
            date=date,
            name=member.display_name,
            user_id=member.id,
            channel_name=after.channel.name,
            join_time=time_str,
            leave_time=""
        )

# Bot を起動
if __name__ == "__main__":
    if not DISCORD_TOKEN:
        print("❌ DISCORD_TOKEN が設定されていません")
    elif not SPREADSHEET_ID:
        print("❌ SPREADSHEET_ID が設定されていません")
    else:
        try:
            print("🚀 Bot を起動しています...")
            bot.run(DISCORD_TOKEN)
        except Exception as e:
            print(f"❌ Bot起動エラー: {e}")
