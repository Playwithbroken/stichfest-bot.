import os
import json
import logging
import asyncio
from datetime import datetime
from typing import List, Dict, Any, Optional

import gspread
from google.oauth2.service_account import Credentials
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup, URLInputFile, ReplyKeyboardMarkup, KeyboardButton
import random

from dotenv import load_dotenv

# --- Load Environment Variables ---
load_dotenv()

# --- Environment Variables ---
BOT_TOKEN = os.getenv("BOT_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDS")
ADMIN_ID = os.getenv("ADMIN_ID") # Optional: Telegram user ID of the admin

# --- Logging ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- States ---
class SetupStates(StatesGroup):
    waiting_for_players = State()

class GameStates(StatesGroup):
    waiting_for_game_type = State()
    waiting_for_winner = State()
    waiting_for_soloist = State()
    waiting_for_re_players = State()
    waiting_for_announcements = State()
    waiting_for_special_points = State()

class AdminStates(StatesGroup):
    waiting_for_rule_value = State()

# --- Google Sheets Setup ---
def get_sheets_client():
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

def get_or_create_daily_sheet(client, spreadsheet_id, players: List[str]):
    sh = client.open_by_key(spreadsheet_id)
    today_str = datetime.now().strftime("%d.%m.%y")
    
    try:
        worksheet = sh.worksheet(today_str)
    except gspread.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=today_str, rows="100", cols="20")
        headers = ["Zeit", "Spiel-Typ", "Gewinner", "Punkte"] + players
        worksheet.append_row(headers)
    
    return worksheet

def get_rules(client, spreadsheet_id):
    sh = client.open_by_key(spreadsheet_id)
    try:
        rules_sheet = sh.worksheet("Rules")
        data = rules_sheet.get_all_records()
        rules = {row['Key']: row['Value'] for row in data}
        return rules
    except gspread.WorksheetNotFound:
        return {
            "SoloMultiplier": 3,
            "Fuchs": 1,
            "Karlchen": 1,
            "Doppelkopf": 1,
            "CentFaktor": 0.05,
            "BasePoint": 1
        }

# --- Scoring Logic ---
def calculate_points(game_data: Dict[str, Any], rules: Dict[str, Any], players: List[str], is_bock: bool = False) -> Dict[str, int]:
    base = int(rules.get("BasePoint", 1))
    
    # Announcements (Ansagen) double the score
    anns = game_data.get("announcements", [])
    multiplier = 2 ** len(anns)
    
    # Extra points (Absagen/Sonderpunkte) add points
    extras = game_data.get("extra_points", []).copy()
    # "Herz-Rundlauf" is a trigger, not an additive point itself in many rules,
    # but the user might want it to count. I'll treat it as a trigger.
    if "Herz-Rundlauf" in extras: extras.remove("Herz-Rundlauf")

    extra_val = len(extras)
    round_points = (base + extra_val) * multiplier
    
    if is_bock:
        round_points *= 2
    
    scores = {p: 0 for p in players}
    
    if game_data["type"] == "Normal":
        re_team = game_data["re_players"]
        kontra_team = [p for p in players if p not in re_team]
        
        if game_data["winner_team"] == "Re":
            for p in re_team: scores[p] = round_points
            for p in kontra_team: scores[p] = -round_points
        else:
            for p in re_team: scores[p] = -round_points
            for p in kontra_team: scores[p] = round_points
            
    elif game_data["type"] == "Solo":
        soloist = game_data["soloist"]
        others = [p for p in players if p != soloist]
        solo_mult = int(rules.get("SoloMultiplier", 3))
        
        if game_data["winner_team"] == "Soloist":
            scores[soloist] = round_points * solo_mult
            for p in others: scores[p] = -round_points
        else:
            scores[soloist] = -(round_points * solo_mult)
            for p in others: scores[p] = round_points
            
    return scores

# --- Bot Initialization ---
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# --- Persistence Helpers ---
def get_bock_count(client, spreadsheet_id):
    sh = client.open_by_key(spreadsheet_id)
    try:
        dashboard = sh.worksheet("Dashboard")
        val = dashboard.acell('B7').value
        return int(val) if val and val.isdigit() else 0
    except:
        return 0

def set_bock_count(client, spreadsheet_id, count):
    sh = client.open_by_key(spreadsheet_id)
    try:
        dashboard = sh.worksheet("Dashboard")
        dashboard.update_acell('B7', count)
    except:
        pass

def get_players_from_dashboard(client, spreadsheet_id):
    sh = client.open_by_key(spreadsheet_id)
    try:
        dashboard = sh.worksheet("Dashboard")
        players = dashboard.col_values(1)[1:] 
        return [p for p in players if p]
    except gspread.WorksheetNotFound:
        return []

def get_main_menu():
    kb = [
        [KeyboardButton(text="🃏 Spiel eintragen"), KeyboardButton(text="📊 Statistik")],
        [KeyboardButton(text="💶 Kasse"), KeyboardButton(text="📜 Regeln")],
        [KeyboardButton(text="🛠 Admin"), KeyboardButton(text="🎲 Mischen")],
        [KeyboardButton(text="💰 Tages-Abrechnung")]
    ]
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)

def format_rule_name(key: str) -> str:
    mapping = {
        "SoloMultiplier": "Solo-Multiplikator (x3, x4...)",
        "Fuchs": "Punkte für Fuchs fangen",
        "Karlchen": "Punkte für Karlchen",
        "Doppelkopf": "Punkte für Doppelkopf",
        "CentFaktor": "Euro pro Punkt (z.B. 0.05)",
        "BasePoint": "Basispunkte pro Spiel"
    }
    return mapping.get(key, key)

def update_dashboard(client, spreadsheet_id, players: List[str]):
    sh = client.open_by_key(spreadsheet_id)
    try:
        dashboard = sh.worksheet("Dashboard")
    except gspread.WorksheetNotFound:
        dashboard = sh.add_worksheet(title="Dashboard", rows="50", cols="10")
    
    # Calculate All-Time Stats
    totals = {p: 0 for p in players}
    games_count = {p: 0 for p in players}
    wins = {p: 0 for p in players}
    
    for ws in sh.worksheets():
        if ws.title in ["Dashboard", "Rules"]: continue
        try:
            data = ws.get_all_records()
            for row in data:
                for p in players:
                    pts = int(row.get(p, 0))
                    totals[p] += pts
                    if pts != 0:
                        games_count[p] += 1
                        if pts > 0: wins[p] += 1
        except: continue

    mvp = max(totals, key=totals.get) if totals else "-"
    pechvogel = min(totals, key=totals.get) if totals else "-"

    # Prepare Dashboard Content
    header = [['Spieler', 'Gesamtpunkte', 'Win-Rate (%)', 'Spiele']]
    rows = []
    for p in players:
        wr = (wins[p] / games_count[p] * 100) if games_count[p] > 0 else 0
        rows.append([p, totals[p], f"{wr:.1f}", games_count[p]])
    
    # Write to Sheet
    dashboard.clear()
    dashboard.update('A1', header + rows)
    
    # Add Highlights Section
    start_row = len(rows) + 3
    dashboard.update(f'A{start_row}', [
        ['🏆 Highlights', ''],
        ['🥇 MVP', mvp],
        ['📉 Pechvogel', pechvogel],
        ['', ''],
        ['🎰 Bock-Zähler', get_bock_count(client, spreadsheet_id)]
    ])
    
    # Basic Formatting (Async in spirit, but synchronous via gspread API)
    try:
        dashboard.format("A1:D1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}})
        dashboard.format(f"A{start_row}", {"textFormat": {"bold": True, "fontSize": 12}})
    except: pass

# --- Handlers ---

@dp.message(F.text == "🃏 Spiel eintragen")
async def menu_score(message: types.Message, state: FSMContext):
    await cmd_score(message, state)

@dp.message(F.text == "📊 Statistik")
async def menu_stats(message: types.Message):
    await cmd_stats(message)

@dp.message(F.text == "💶 Kasse")
async def menu_kasse(message: types.Message):
    await cmd_kasse(message)

@dp.message(F.text == "📜 Regeln")
async def menu_rules(message: types.Message):
    await cmd_rules(message)

@dp.message(F.text == "🛠 Admin")
async def menu_admin(message: types.Message):
    await cmd_admin(message)

@dp.message(F.text == "🎲 Mischen")
async def menu_mischen(message: types.Message):
    await cmd_mischen(message)

@dp.message(F.text == "💰 Tages-Abrechnung")
async def menu_settlement(message: types.Message):
    await cmd_settlement(message)

@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    await message.answer(
        "Willkommen beim **Stichfest & Saufbereit** Bot! 🃏🍻\n\nNutze das Menü unten für die schnelle Bedienung.",
        reply_markup=get_main_menu(),
        parse_mode="Markdown"
    )
    # Check if players exist
    try:
        client = get_sheets_client()
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        if not players:
            await message.answer("Es sind noch keine Spieler registriert. Bitte gib die Namen der 4 oder 5 Spieler ein (kommagetrennt):")
            await state.set_state(SetupStates.waiting_for_players)
    except Exception as e:
        logger.error(f"Start error: {e}")

@dp.message(SetupStates.waiting_for_players)
async def process_players(message: types.Message, state: FSMContext):
    players = [p.strip() for p in message.text.split(",") if p.strip()]
    if len(players) not in [4, 5]:
        await message.answer("Bitte gib genau 4 oder 5 Spieler an.")
        return
    
    await state.update_data(players=players)
    try:
        client = get_sheets_client()
        update_dashboard(client, SPREADSHEET_ID, players)
        await message.answer(f"Spieler registriert: {', '.join(players)}\n\nAlle Statistiken werden ab jetzt auf dem Live-Dashboard getrackt! 📊")
    except Exception as e:
        logger.error(f"Error updating dashboard: {e}")
        await message.answer(f"Fehler beim Speichern in Google Sheets: {e}")
    await state.clear()

@dp.message(Command("score"))
async def cmd_score(message: types.Message, state: FSMContext):
    try:
        client = get_sheets_client()
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        if not players:
            await message.answer("Keine Spieler gefunden. Bitte nutze /start.")
            return
        await state.update_data(players=players)
    except Exception as e:
        await message.answer(f"Fehler beim Laden der Spieler: {e}")
        return
    
    kb = InlineKeyboardBuilder()
    kb.button(text="Normal 🃏", callback_data="type:Normal")
    kb.button(text="Solo 👤", callback_data="type:Solo")
    kb.adjust(1)
    await message.answer("Was für ein Spiel war es?", reply_markup=kb.as_markup())
    await state.set_state(GameStates.waiting_for_game_type)

@dp.callback_query(F.data.startswith("type:"))
async def process_game_type(callback: types.CallbackQuery, state: FSMContext):
    game_type = callback.data.split(":")[1]
    await state.update_data(type=game_type)
    data = await state.get_data()
    players = data["players"]

    if game_type == "Normal":
        await state.update_data(re_players=[]) 
        kb = InlineKeyboardBuilder()
        for p in players:
            kb.button(text=f"⬜ {p}", callback_data=f"toggle_re:{p}")
        kb.adjust(2)
        await callback.message.edit_text("Wer ist Team Re? (Wähle 2 Spieler)", reply_markup=kb.as_markup())
        await state.set_state(GameStates.waiting_for_re_players)
    else:
        kb = InlineKeyboardBuilder()
        for p in players:
            kb.button(text=p, callback_data=f"soloist:{p}")
        await callback.message.edit_text("Wer war der Solist?", reply_markup=kb.as_markup())
        await state.set_state(GameStates.waiting_for_soloist)

@dp.callback_query(F.data.startswith("toggle_re:"))
async def handle_re_selection(callback: types.CallbackQuery, state: FSMContext):
    p_selected = callback.data.split(":")[1]
    data = await state.get_data()
    re_players = data.get("re_players", [])
    players = data["players"]
    
    if p_selected in re_players:
        re_players.remove(p_selected)
    else:
        if len(re_players) < 2:
            re_players.append(p_selected)
    
    await state.update_data(re_players=re_players)
    kb = InlineKeyboardBuilder()
    for p in players:
        prefix = "✅ " if p in re_players else "⬜ "
        kb.button(text=f"{prefix}{p}", callback_data=f"toggle_re:{p}")
    kb.adjust(2)
    if len(re_players) == 2:
        kb.row(InlineKeyboardButton(text="Bestätigen ✅", callback_data="re_confirmed"))
    await callback.message.edit_reply_markup(reply_markup=kb.as_markup())

@dp.callback_query(F.data == "re_confirmed")
async def confirm_re_team(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    re_str = ", ".join(data["re_players"])
    kb = InlineKeyboardBuilder()
    kb.button(text="Team Re 🎉", callback_data="winner:Re")
    kb.button(text="Team Kontra 👊", callback_data="winner:Kontra")
    kb.adjust(1)
    await callback.message.edit_text(f"Team Re: {re_str}\n\nWer hat gewonnen?", reply_markup=kb.as_markup())
    await state.set_state(GameStates.waiting_for_winner)

@dp.callback_query(F.data.startswith("soloist:"))
async def process_soloist(callback: types.CallbackQuery, state: FSMContext):
    soloist = callback.data.split(":")[1]
    await state.update_data(soloist=soloist)
    kb = InlineKeyboardBuilder()
    kb.button(text="Soloist gewonnen 🏆", callback_data="winner:Soloist")
    kb.button(text="Gegenpartei gewonnen 💥", callback_data="winner:Others")
    kb.adjust(1)
    await callback.message.edit_text(f"Hat {soloist} gewonnen?", reply_markup=kb.as_markup())
    await state.set_state(GameStates.waiting_for_winner)

@dp.callback_query(F.data.startswith("winner:"))
async def handle_winner_selection(callback: types.CallbackQuery, state: FSMContext):
    winner = callback.data.split(":")[1]
    await state.update_data(winner_team=winner, announcements=[])
    kb = InlineKeyboardBuilder()
    options = ["Re", "Kontra", "Keine 90", "Keine 60", "Keine 30"]
    for opt in options:
        kb.button(text=f"⬜ {opt}", callback_data=f"toggle_ann:{opt}")
    kb.adjust(2)
    kb.row(InlineKeyboardButton(text="Weiter ➡️", callback_data="ann_done"))
    await callback.message.edit_text("Welche Ansagen wurden gemacht?", reply_markup=kb.as_markup())
    await state.set_state(GameStates.waiting_for_announcements)

@dp.callback_query(F.data.startswith("toggle_ann:"))
async def handle_announcement_toggle(callback: types.CallbackQuery, state: FSMContext):
    opt = callback.data.split(":")[1]
    data = await state.get_data()
    anns = data.get("announcements", [])
    if opt in anns: anns.remove(opt)
    else: anns.append(opt)
    await state.update_data(announcements=anns)
    kb = InlineKeyboardBuilder()
    options = ["Re", "Kontra", "Keine 90", "Keine 60", "Keine 30"]
    for o in options:
        prefix = "✅ " if o in anns else "⬜ "
        kb.button(text=f"{prefix}{o}", callback_data=f"toggle_ann:{o}")
    kb.adjust(2)
    kb.row(InlineKeyboardButton(text="Weiter ➡️", callback_data="ann_done"))
    await callback.message.edit_reply_markup(reply_markup=kb.as_markup())

@dp.callback_query(F.data == "ann_done")
async def handle_announcement_done(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(extra_points=[])
    kb = InlineKeyboardBuilder()
    options = ["Fuchs", "Karlchen", "Doppelkopf", "Keine 90", "Keine 60", "Keine 30", "Schwarz", "Herz-Rundlauf"]
    for opt in options:
        kb.button(text=f"⬜ {opt}", callback_data=f"toggle_extra:{opt}")
    kb.adjust(2)
    kb.row(InlineKeyboardButton(text="Abschließen 🏁", callback_data="extra_done"))
    await callback.message.edit_text("Welche Sonderpunkte/Absagen gab es?", reply_markup=kb.as_markup())
    await state.set_state(GameStates.waiting_for_special_points)

@dp.callback_query(F.data.startswith("toggle_extra:"))
async def handle_extra_toggle(callback: types.CallbackQuery, state: FSMContext):
    opt = callback.data.split(":")[1]
    data = await state.get_data()
    extras = data.get("extra_points", [])
    if opt in extras: extras.remove(opt)
    else: extras.append(opt)
    await state.update_data(extra_points=extras)
    kb = InlineKeyboardBuilder()
    options = ["Fuchs", "Karlchen", "Doppelkopf", "Keine 90", "Keine 60", "Keine 30", "Schwarz", "Herz-Rundlauf"]
    for o in options:
        prefix = "✅ " if o in extras else "⬜ "
        kb.button(text=f"{prefix}{o}", callback_data=f"toggle_extra:{o}")
    kb.adjust(2)
    kb.row(InlineKeyboardButton(text="Abschließen 🏁", callback_data="extra_done"))
    await callback.message.edit_reply_markup(reply_markup=kb.as_markup())

@dp.callback_query(F.data == "extra_done")
async def handle_final_score(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    if "players" not in data:
        client = get_sheets_client()
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        await state.update_data(players=players)
        data = await state.get_data()
    players = data["players"]
    
    try:
        await callback.message.edit_text("Berechne Punkte... ⏳")
        client = get_sheets_client()
        rules = get_rules(client, SPREADSHEET_ID)
        
        # Bock-Logik
        current_bock = get_bock_count(client, SPREADSHEET_ID)
        is_bock_round = current_bock > 0
        
        if data["type"] == "Normal" and ("re_players" not in data or len(data["re_players"]) != 2):
            await callback.message.answer("⚠️ Fehler: Team Re wurde nicht korrekt festgelegt.")
            await state.clear()
            return

        scores = calculate_points(data, rules, players, is_bock=is_bock_round)
        
        # Bock-Zähler aktualisieren
        new_bock = current_bock
        if is_bock_round: new_bock -= 1
        if "Herz-Rundlauf" in data.get("extra_points", []):
            new_bock += 4
        
        if new_bock != current_bock:
            set_bock_count(client, SPREADSHEET_ID, new_bock)

        # Log to Sheet
        sheet = get_or_create_daily_sheet(client, SPREADSHEET_ID, players)
        row = [datetime.now().strftime("%H:%M:%S"), data["type"], data["winner_team"], sum([s for s in scores.values() if s > 0])]
        for p in players: row.append(scores[p])
        sheet.append_row(row)
        
        # Success Message
        score_details = "\n".join([f"• {p}: `{s:+}` Pkt" for p, s in scores.items()])
        summary = f"**Runde geloggt! ✅**\n\n"
        
        if is_bock_round: 
            summary += "🔥 🃏 **BOCKRUNDE (Doppelte Punkte!)** 🃏 🔥\n"
        
        summary += f"📍 Typ: *{data['type']}* | Sieger: *{data['winner_team']}*\n"
        
        if data.get('announcements'): 
            summary += f"📢 Ansagen: {', '.join([f'_{a}_' for a in data['announcements']])}\n"
        if data.get('extra_points'): 
            summary += f"✨ Extras: {', '.join([f'_{e}_' for e in data['extra_points']])}\n"
        
        summary += f"\n{score_details}\n"
        
        if new_bock > 0: 
            summary += f"\n🎰 **Noch {new_bock} Bockrunden verbleibend!**"
        elif is_bock_round and new_bock == 0: 
            summary += "\n🏁 Die Bockrunden sind vorbei. Ab jetzt wieder normal!"
        
        if "Herz-Rundlauf" in data.get("extra_points", []):
            summary += "\n\n📢 **HERZ-RUNDLAUF!** Das gibt 4 neue Bockrunden! 💥"
            
        await callback.message.edit_text(summary, parse_mode="Markdown")
        
        # Proactively update dashboard with new stats
        update_dashboard(client, SPREADSHEET_ID, players)
        
    except Exception as e:
        logger.error(f"Scoring error: {e}")
        await callback.message.answer(f"❌ Fehler beim Loggen: {e}")
    await state.clear()

@dp.message(Command("kasse"))
async def cmd_kasse(message: types.Message):
    try:
        client = get_sheets_client()
        rules = get_rules(client, SPREADSHEET_ID)
        cent_faktor = float(rules.get("CentFaktor", 0.05))
        sh = client.open_by_key(SPREADSHEET_ID)
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        totals = {p: 0 for p in players}
        for ws in sh.worksheets():
            if ws.title in ["Dashboard", "Rules"]: continue
            data = ws.get_all_records()
            for row in data:
                for p in players: totals[p] += int(row.get(p, 0))
        res = "💶 **Aktueller Kassenstand:**\n"
        for p, s in totals.items():
            euro = s * cent_faktor
            res += f"{p}: {s} Pkt ({euro:+.2f}€)\n"
        await message.answer(res, parse_mode="Markdown")
    except Exception as e:
        await message.answer(f"Fehler: {e}")

@dp.message(Command("dashboard"))
async def cmd_dashboard(message: types.Message):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    kb = InlineKeyboardBuilder()
    kb.button(text="Zum Google Sheet 📊", url=url)
    await message.answer("Hier ist der Link zum Live-Dashboard:", reply_markup=kb.as_markup())

@dp.message(Command("undo"))
async def cmd_undo(message: types.Message):
    try:
        client = get_sheets_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        today_str = datetime.now().strftime("%d.%m.%y")
        try:
            worksheet = sh.worksheet(today_str)
            rows = worksheet.get_all_values()
            if len(rows) <= 1:
                await message.answer("Keine Runden zum Rückgängigmachen vorhanden.")
                return
            
            # Get last row values to check for Bock adjustments
            last_row = rows[-1]
            # Since we don't store "was it bock" in the row explicitly in a way that's easy to reverse,
            # this is a bit tricky. But the last entry in main.py logic subtracts bock if is_bock_round.
            
            # Simple delete for now
            worksheet.delete_rows(len(rows))
            await message.answer("Letzte Runde wurde erfolgreich gelöscht! 🗑️")
        except gspread.WorksheetNotFound:
            await message.answer("Heute wurden noch keine Runden gespielt.")
    except Exception as e:
        await message.answer(f"Fehler beim Undo: {e}")

@dp.message(Command("mischen"))
async def cmd_mischen(message: types.Message):
    try:
        client = get_sheets_client()
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        if not players:
            await message.answer("Keine Spieler gefunden. Nutze /start.")
            return
        
        shuffled = players.copy()
        random.shuffle(shuffled)
        
        res = "🎲 **Neue Sitzordnung:**\n\n"
        for i, p in enumerate(shuffled, 1):
            res += f"{i}. {p}\n"
        res += "\nDer Erste gibt an! 🃏"
        await message.answer(res, parse_mode="Markdown")
    except Exception as e:
        await message.answer(f"Fehler beim Mischen: {e}")

@dp.message(Command("stats"))
async def cmd_stats(message: types.Message):
    try:
        await message.answer("Berechne Statistiken... 📊")
        client = get_sheets_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        totals = {p: 0 for p in players}
        games_count = {p: 0 for p in players}
        wins = {p: 0 for p in players}
        
        for ws in sh.worksheets():
            if ws.title in ["Dashboard", "Rules"]: continue
            data = ws.get_all_records()
            for row in data:
                for p in players:
                    pts = int(row.get(p, 0))
                    totals[p] += pts
                    if pts != 0:
                        games_count[p] += 1
                        if pts > 0: wins[p] += 1
        
        # Determine MVP (Highest Total) and Pechvogel (Lowest Total)
        mvp = max(totals, key=totals.get)
        pechvogel = min(totals, key=totals.get)
        
        res = "🏆 **Stichfest-Statistiken** 🏆\n\n"
        for p in players:
            win_rate = (wins[p] / games_count[p] * 100) if games_count[p] > 0 else 0
            res += f"👤 *{p}*:\n   - Pkt: {totals[p]}\n   - Win-Rate: {win_rate:.1f}%\n"
        
        res += f"\n🥇 **MVP:** {mvp} ({totals[mvp]} Pkt)\n"
        res += f"📉 **Pechvogel:** {pechvogel} ({totals[pechvogel]} Pkt)\n"
        await message.answer(res, parse_mode="Markdown")
    except Exception as e:
        await message.answer(f"Fehler bei Stats: {e}")

@dp.message(Command("settlement"))
async def cmd_settlement(message: types.Message):
    try:
        await message.answer("Erstelle Tages-Abrechnung... 💰")
        client = get_sheets_client()
        rules = get_rules(client, SPREADSHEET_ID)
        cent_faktor = float(rules.get("CentFaktor", 0.05))
        sh = client.open_by_key(SPREADSHEET_ID)
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        
        today_str = datetime.now().strftime("%d.%m.%y")
        try:
            ws = sh.worksheet(today_str)
            data = ws.get_all_records()
        except gspread.WorksheetNotFound:
            await message.answer("Heute wurden noch keine Runden gespielt. Nichts abzurechnen! 🍻")
            return

        today_totals = {p: 0 for p in players}
        for row in data:
            for p in players:
                today_totals[p] += int(row.get(p, 0))
        
        res = f"💰 **Abrechnung für heute ({today_str}):**\n\n"
        for p, s in today_totals.items():
            euro = s * cent_faktor
            status = "zahlt" if euro < 0 else "bekommt"
            res += f"👤 **{p}**: {s} Pkt -> {abs(euro):.2f}€ {status}\n"
        
        res += "\nProst! 🍻"
        await message.answer(res, parse_mode="Markdown")
    except Exception as e:
        await message.answer(f"Fehler bei Abrechnung: {e}")

@dp.message(Command("me"))
async def cmd_me(message: types.Message):
    # Try to match Telegram name with registered player names
    tg_name = message.from_user.full_name
    try:
        client = get_sheets_client()
        players = get_players_from_dashboard(client, SPREADSHEET_ID)
        
        # Simple fuzzy match (if TG name is in registered players)
        match = None
        for p in players:
            if p.lower() in tg_name.lower() or tg_name.lower() in p.lower():
                match = p
                break
        
        if not match:
            await message.answer(f"Ich konnte dich nicht automatisch zuordnen (Telegram: {tg_name}).\nRegistrierte Spieler: {', '.join(players)}")
            return
            
        # Aggregate logic same as above but just for one player
        sh = client.open_by_key(SPREADSHEET_ID)
        total = 0
        games = 0
        w = 0
        for ws in sh.worksheets():
            if ws.title in ["Dashboard", "Rules"]: continue
            data = ws.get_all_records()
            for row in data:
                pts = int(row.get(match, 0))
                total += pts
                if pts != 0:
                    games += 1
                    if pts > 0: w += 1
        
        win_rate = (w / games * 100) if games > 0 else 0
        res = f"🎴 **Deine Statistik ({match})** 🎴\n\n"
        res += f"• Gesamtpunkte: {total}\n"
        res += f"• Spiele: {games}\n"
        res += f"• Siege: {w}\n"
        res += f"• Win-Rate: {win_rate:.1f}%\n"
        
        if total > 0: res += "\nLäuft bei dir! 🎉"
        else: res += "\nDa ist noch Luft nach oben... 🍻"
        
        await message.answer(res, parse_mode="Markdown")
    except Exception as e:
        await message.answer(f"Fehler: {e}")

@dp.message(Command("rules"))
async def cmd_rules(message: types.Message):
    try:
        client = get_sheets_client()
        rules = get_rules(client, SPREADSHEET_ID)
        res = "📜 **Aktuelle Spielregeln:**\n\n"
        for k, v in rules.items():
            label = format_rule_name(k)
            res += f"• **{label}**: `{v}`\n"
        
        url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        kb = InlineKeyboardBuilder()
        kb.button(text="Im Sheet ansehen ↗️", url=url)
        
        await message.answer(res, parse_mode="Markdown", reply_markup=kb.as_markup())
    except Exception as e:
        await message.answer(f"Fehler: {e}")

@dp.message(Command("admin"))
async def cmd_admin(message: types.Message):
    user_id = str(message.from_user.id)
    if ADMIN_ID and user_id != ADMIN_ID:
        await message.answer(f"🚫 Zugriff verweigert. Deine ID ({user_id}) ist nicht als Admin hinterlegt.")
        return
    
    kb = InlineKeyboardBuilder()
    kb.button(text="Spieler zurücksetzen 👥", callback_data="admin_reset_players")
    kb.button(text="Bockrunden löschen 🎰", callback_data="admin_reset_bock")
    kb.button(text="Regeln anpassen ⚙️", callback_data="admin_edit_rules")
    kb.button(text="Einladungs-Text 📩", callback_data="admin_invite")
    kb.adjust(1)
    await message.answer("🛠 **Admin Panel**\nWas möchtest du tun?", reply_markup=kb.as_markup())

@dp.callback_query(F.data == "admin_reset_players")
async def handle_admin_reset_players(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.edit_text("⚠️ Bist du sicher? Dies löscht die Spieler-Zuordnung (nicht die Punkte im Sheet).",
                                    reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                                        [InlineKeyboardButton(text="Ja, Reset!", callback_data="admin_confirm_reset")],
                                        [InlineKeyboardButton(text="Abbrechen", callback_data="admin_cancel")]
                                    ]))

@dp.callback_query(F.data == "admin_confirm_reset")
async def handle_confirm_reset(callback: types.CallbackQuery):
    try:
        client = get_sheets_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        dashboard = sh.worksheet("Dashboard")
        # Clear players column
        dashboard.update('A2:A10', [[''] for _ in range(9)])
        await callback.message.edit_text("✅ Spieler-Zuordnung wurde zurückgesetzt. Nutze /start für ein neues Setup.")
    except Exception as e:
        await callback.message.answer(f"Fehler: {e}")

@dp.callback_query(F.data == "admin_reset_bock")
async def handle_reset_bock(callback: types.CallbackQuery):
    try:
        client = get_sheets_client()
        set_bock_count(client, SPREADSHEET_ID, 0)
        await callback.message.edit_text("✅ Bock-Runden wurden auf 0 gesetzt.")
    except Exception as e:
        await callback.message.answer(f"Fehler: {e}")

@dp.callback_query(F.data == "admin_invite")
async def handle_admin_invite(callback: types.CallbackQuery):
    bot_info = await bot.get_me()
    invite_text = (
        f"🃏 **Einladung zur Doppelkopf-Runde!** 🃏\n\n"
        f"Tretet dem Bot bei, um Punkte zu tracken und Statistiken zu sehen:\n\n"
        f"👉 [t.me/{bot_info.username}](t.me/{bot_info.username})\n\n"
        f"Viel Erfolg beim Stichfest werden! 🍻"
    )
    await callback.message.edit_text(f"Kopiere diesen Text für deine Freunde:\n\n`{invite_text}`", parse_mode="Markdown")

@dp.callback_query(F.data == "admin_edit_rules")
async def handle_admin_edit_rules(callback: types.CallbackQuery):
    try:
        client = get_sheets_client()
        rules = get_rules(client, SPREADSHEET_ID)
        kb = InlineKeyboardBuilder()
        for k in rules.keys():
            label = format_rule_name(k)
            kb.button(text=label, callback_data=f"edit_rule:{k}")
        kb.adjust(1)
        kb.row(InlineKeyboardButton(text="Zurück ⬅️", callback_data="admin_cancel"))
        await callback.message.edit_text("Welche Regel möchtest du ändern?", reply_markup=kb.as_markup())
    except Exception as e:
        await callback.message.answer(f"Fehler: {e}")

@dp.callback_query(F.data.startswith("edit_rule:"))
async def process_edit_rule(callback: types.CallbackQuery, state: FSMContext):
    rule_key = callback.data.split(":")[1]
    await state.update_data(editing_rule=rule_key)
    label = format_rule_name(rule_key)
    await callback.message.edit_text(f"Gib bitte den neuen Wert für **{label}** ein (als Zahl):")
    await state.set_state(AdminStates.waiting_for_rule_value)

@dp.message(AdminStates.waiting_for_rule_value)
async def handle_rule_value_input(message: types.Message, state: FSMContext):
    new_val = message.text.strip().replace(",", ".")
    data = await state.get_data()
    rule_key = data["editing_rule"]
    
    try:
        # Check if it's a valid number
        float(new_val)
        
        client = get_sheets_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        rules_sheet = sh.worksheet("Rules")
        
        # Find the row with the key
        cells = rules_sheet.findall(rule_key)
        if not cells:
            await message.answer(f"❌ Regel '{rule_key}' wurde im Sheet nicht gefunden.")
            await state.clear()
            return
            
        row = cells[0].row
        rules_sheet.update_cell(row, 2, new_val) # Assuming Value is col 2
        
        await message.answer(f"✅ Die Regel **{format_rule_name(rule_key)}** wurde auf `{new_val}` aktualisiert!", reply_markup=get_main_menu())
        await state.clear()
    except ValueError:
        await message.answer("❌ Ungültige Eingabe. Bitte gib eine Zahl ein (z.B. 3 oder 0.05).")
    except Exception as e:
        await message.answer(f"❌ Fehler beim Speichern: {e}")
        await state.clear()

@dp.callback_query(F.data == "admin_cancel")
async def handle_admin_cancel(callback: types.CallbackQuery):
    await callback.message.edit_text("Vorgang abgebrochen.")

async def main():
    logger.info("Bot starting...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
