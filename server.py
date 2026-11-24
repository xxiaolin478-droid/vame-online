import asyncio
import json
import random
import string
from pathlib import Path
from typing import Dict, List, Optional, Set

import pandas as pd
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI()

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
COAPPEAR_FILE = BASE_DIR / "共演表.xlsx"
CHARACTER_FILE = BASE_DIR / "キャラ表2.xlsx"

# --------- 起動時にExcelを1回だけ高速読み込み ---------
print("Excelファイルを読み込んでいます...")

char_df = pd.read_excel(CHARACTER_FILE)
df_full = pd.read_excel(COAPPEAR_FILE, index_col=0)

char_df = char_df.drop(index=char_df.index[0])
first_col = char_df.columns[0]

character_map: Dict[str, Dict[str, List[str]]] = {}
anime_titles_for_char = [str(c) for c in char_df.columns[1:]]
char_actors = [str(a).strip() for a in char_df[first_col]]
char_values = char_df.iloc[:, 1:].to_numpy()

for r_idx, actor in enumerate(char_actors):
    if not actor or actor in ("nan", "NaN", "None"):
        continue
    row_cells = char_values[r_idx]
    for c_idx, anime_title in enumerate(anime_titles_for_char):
        cell = row_cells[c_idx]
        if cell is None or (isinstance(cell, float) and pd.isna(cell)):
            continue
        cell_str = str(cell).strip()
        if not cell_str or cell_str in ("nan", "NaN"):
            continue
        roles = [r.strip() for r in cell_str.split("`") if r.strip()]
        if roles:
            character_map.setdefault(anime_title, {}).setdefault(actor, []).extend(roles)

all_actors: List[str] = [
    a.strip() for a in char_df[first_col].dropna().astype(str).tolist() if a.strip()
]

all_coappear_full: Dict[tuple, List[str]] = {}
unique_occurrences_full = []
coactor_count_map: Dict[tuple, int] = {}
anime_max_coactor_map: Dict[str, int] = {}
_temp_coactors: Dict[tuple, Set[str]] = {}

idx_list = [str(x).strip() for x in df_full.index]
col_list = [str(x).strip() for x in df_full.columns]
coappear_values = df_full.to_numpy()

for i, a1 in enumerate(idx_list):
    row_vals = coappear_values[i]
    for j, a2 in enumerate(col_list):
        cell = row_vals[j]
        if cell is None or (isinstance(cell, float) and pd.isna(cell)):
            continue
        cell_str = str(cell).strip()
        if cell_str in ("-", "", "nan", "NaN"):
            continue
        animes = [x.strip() for x in cell_str.split("`") if x.strip()]
        if not animes:
            continue

        key_dir = (a1, a2)
        if key_dir not in all_coappear_full:
            all_coappear_full[key_dir] = animes
        else:
            all_coappear_full[key_dir] = list(set(all_coappear_full[key_dir] + animes))

        s_key = (a1, a2) if a1 <= a2 else (a2, a1)
        for anime in animes:
            unique_occurrences_full.append((s_key[0], s_key[1], anime))

unique_occurrences_full = list(set(unique_occurrences_full))

for a1, a2, anime in unique_occurrences_full:
    _temp_coactors.setdefault((anime, a1), set()).add(a2)
    _temp_coactors.setdefault((anime, a2), set()).add(a1)

for (anime, actor), coactors in _temp_coactors.items():
    cnt = len(coactors)
    coactor_count_map[(anime, actor)] = cnt
    if cnt > anime_max_coactor_map.get(anime, 0):
        anime_max_coactor_map[anime] = cnt

for anime in anime_max_coactor_map:
    if anime_max_coactor_map[anime] <= 0:
        anime_max_coactor_map[anime] = 1

print(f"読み込み完了！ (全声優数: {len(all_actors)}名)")


def get_first_character(anime: str, actor: str) -> str:
    roles = character_map.get(anime, {}).get(actor, [])
    return roles[0] if roles else "？"


# --------- グローバル接続ユーザー管理 ---------
connected_users: Dict[str, dict] = {}


def get_user_status(u: dict) -> str:
    rid = u.get("room_id")
    if not rid or rid not in rooms:
        return "ホーム画面"
    room = rooms[rid]
    if room.phase == "LOBBY":
        return f"ルーム {rid}（待機中）"
    elif room.phase == "PLAYING":
        return f"ルーム {rid}（対戦中）"
    else:
        return f"ルーム {rid}（結果画面）"


async def broadcast_online_users():
    user_list = [
        {
            "id": u["id"],
            "name": u["name"],
            "avatar": u.get("avatar", ""),
            "status": get_user_status(u),
        }
        for u in connected_users.values()
    ]
    payload = json.dumps({"type": "ONLINE_USERS", "users": user_list}, ensure_ascii=False)
    for u in list(connected_users.values()):
        ws = u.get("ws")
        if ws:
            try:
                await ws.send_text(payload)
            except Exception:
                pass


async def broadcast_user_login_toast(login_client_id: str, name: str, avatar: str):
    payload = json.dumps(
        {
            "type": "USER_LOGIN_NOTIFY",
            "clientId": login_client_id,
            "name": name,
            "avatar": avatar,
        },
        ensure_ascii=False,
    )
    for cid, u in list(connected_users.items()):
        if cid == login_client_id:
            continue
        ws = u.get("ws")
        if ws:
            try:
                await ws.send_text(payload)
            except Exception:
                pass


# --------- ルーム管理クラス ---------
class GameRoom:
    def __init__(self, room_id: str, host_id: str, settings: dict):
        self.room_id = room_id
        self.host_id = host_id
        self.settings = settings
        self.phase = "LOBBY"  # LOBBY | PLAYING | RESULT
        self.connections: Dict[str, WebSocket] = {}
        self.lobby_humans: List[dict] = []
        self.ai_configs: List[dict] = settings.get("aiList", [])
        self.match_count = 0

        self.players: List[dict] = []
        self.voice_actors: List[str] = []
        self.deck: List[str] = []
        self.deck_refill_count = 0
        self.max_refill = 10
        self.field: List[str] = []
        self.turn = 1
        self.current_pid = 0
        self.acted_in_round = False
        self.used_anime: Set[str] = set()
        self.logs: List[str] = []
        self.ai_known_sets: Dict[int, Dict[tuple, Set[str]]] = {}

    def add_log(self, msg: str):
        self.logs.append(msg)
        if len(self.logs) > 60:
            self.logs.pop(0)

    def draw_card(self) -> Optional[str]:
        if not self.deck:
            if self.deck_refill_count < self.max_refill:
                self.deck = self.voice_actors.copy()
                random.shuffle(self.deck)
                self.deck_refill_count += 1
                self.add_log(f"山札をリフル ({self.deck_refill_count}/{self.max_refill})")
            else:
                return None
        return self.deck.pop()

    def is_deck_completely_empty(self) -> bool:
        return len(self.deck) == 0 and self.deck_refill_count >= self.max_refill

    def find_common_anime(self, v1: str, v2: str, pid: int) -> List[str]:
        if pid < 0 or pid >= len(self.players):
            return []
        player = self.players[pid]
        if not player["isAI"]:
            res = set()
            res.update(all_coappear_full.get((v1, v2), []))
            res.update(all_coappear_full.get((v2, v1), []))
            return sorted(list(res))
        else:
            known = self.ai_known_sets.get(pid, {})
            res = set()
            res.update(known.get((v1, v2), set()))
            res.update(known.get((v2, v1), set()))
            return sorted(list(res))

    def start_game(self):
        self.match_count += 1
        limit = max(4, min(int(self.settings.get("actorLimit", len(all_actors))), len(all_actors)))
        self.voice_actors = all_actors[:limit]

        for h in self.lobby_humans:
            if h["id"] in connected_users:
                h["name"] = connected_users[h["id"]]["name"]
                h["avatar"] = connected_users[h["id"]].get("avatar", "")

        humans = [dict(h) for h in self.lobby_humans]
        random.shuffle(humans)

        self.players = []
        for h in humans:
            self.players.append({
                "id": h["id"],
                "name": h["name"],
                "avatar": h.get("avatar", ""),
                "isAI": False,
                "aiPct": 100,
                "hand": [],
                "usedAnime": [],
                "finished": False,
                "rank": None,
                "leftRoom": False,
            })

        for idx, ai in enumerate(self.ai_configs):
            self.players.append({
                "id": f"ai_{idx+1}",
                "name": ai.get("name", f"AI-{idx+1}"),
                "avatar": "",
                "isAI": True,
                "aiPct": float(ai.get("pct", 80)),
                "hand": [],
                "usedAnime": [],
                "finished": False,
                "rank": None,
                "leftRoom": False,
            })

        self.ai_known_sets = {}
        va_set = set(self.voice_actors)
        for pid, p in enumerate(self.players):
            if not p["isAI"]:
                continue
            pct = p["aiPct"]
            known_map: Dict[tuple, Set[str]] = {}
            for a1, a2, anime in unique_occurrences_full:
                if a1 not in va_set or a2 not in va_set:
                    continue
                c1 = coactor_count_map.get((anime, a1), 0)
                c2 = coactor_count_map.get((anime, a2), 0)
                mc = anime_max_coactor_map.get(anime, 1)
                b1 = 0.5 + 0.5 * ((c1 / mc) ** 2)
                b2 = 0.5 + 0.5 * ((c2 / mc) ** 2)
                p_anime = pct * ((b1 + b2) / 2)
                if random.random() * 100 < p_anime:
                    known_map.setdefault((a1, a2), set()).add(anime)
                    known_map.setdefault((a2, a1), set()).add(anime)
            self.ai_known_sets[pid] = known_map

        self.deck = self.voice_actors.copy()
        random.shuffle(self.deck)
        self.deck_refill_count = 0
        self.used_anime = set()
        self.logs = []
        self.turn = 1
        self.current_pid = 0
        self.acted_in_round = False

        start_hand = int(self.settings.get("startHandSize", 5))
        for p in self.players:
            p["hand"] = [c for _ in range(start_hand) if (c := self.draw_card())]

        self.field = [c for _ in range(3) if (c := self.draw_card())]
        self.phase = "PLAYING"
        self.add_log(f"第{self.match_count}戦 開始！人間プレイヤーの順番がシャッフルされました。")

    def get_state_for_client(self, client_id: str) -> dict:
        my_pid = next((i for i, p in enumerate(self.players) if p["id"] == client_id), -1)
        public_players = []
        for i, p in enumerate(self.players):
            public_players.append({
                "pid": i,
                "id": p["id"],
                "name": p["name"],
                "avatar": p.get("avatar", ""),
                "isAI": p["isAI"],
                "aiPct": p["aiPct"],
                "handCount": len(p["hand"]),
                "hand": p["hand"] if i == my_pid else [],
                "usedAnime": p["usedAnime"],
                "finished": p["finished"],
                "rank": p["rank"],
                "leftRoom": p.get("leftRoom", False),
            })

        return {
            "type": "STATE",
            "roomId": self.room_id,
            "hostId": self.host_id,
            "phase": self.phase,
            "matchCount": self.match_count,
            "settings": self.settings,
            "lobbyHumans": self.lobby_humans,
            "aiConfigs": self.ai_configs,
            "myPid": my_pid,
            "turn": self.turn,
            "currentPid": self.current_pid,
            "field": self.field,
            "deckCount": len(self.deck),
            "refillCount": self.deck_refill_count,
            "maxRefill": self.max_refill,
            "players": public_players,
            "logs": self.logs,
        }

    async def broadcast(self):
        for cid, ws in list(self.connections.items()):
            try:
                await ws.send_text(json.dumps(self.get_state_for_client(cid), ensure_ascii=False))
            except Exception:
                pass
        await broadcast_online_users()

    def finish_game_and_assign_ranks(self, reason_log: str):
        active = [x for x in self.players if not x["finished"]]
        for rem in active:
            rem["finished"] = True
            rem["rank"] = sum(1 for x in self.players if x["finished"])
        self.phase = "RESULT"
        self.add_log(reason_log)

    def check_victory_and_advance(self):
        endless = bool(self.settings.get("endlessMode", False))
        start_hand = int(self.settings.get("startHandSize", 5))
        extra_card = int(self.settings.get("extraCard", 3))
        end_limit = start_hand + extra_card
        n = len(self.players)

        p = self.players[self.current_pid]
        if not endless and len(p["hand"]) == 0 and not p["finished"]:
            p["finished"] = True
            p["rank"] = sum(1 for x in self.players if x["finished"])
            self.add_log(f"🏆 {p['name']} が手札を使い切り {p['rank']}位 で上がり！")

        if endless:
            drawn = self.draw_card()
            if drawn:
                p["hand"].append(drawn)
            if len(p["hand"]) >= end_limit and not p["finished"]:
                p["finished"] = True
                p["rank"] = sum(1 for x in self.players if x["finished"])
                self.add_log(f"💀 上限 {end_limit} 枚到達: {p['name']} が脱落 ({p['rank']}位)")

            if self.is_deck_completely_empty():
                self.finish_game_and_assign_ranks("📭 山札（リフル上限含む）がなくなったためゲーム終了！")
                return

        active = [x for x in self.players if not x["finished"]]
        humans_active = [x for x in active if not x["isAI"]]
        if (n == 1 and len(active) == 0) or (n > 1 and (len(active) <= 1 or len(humans_active) == 0)):
            self.finish_game_and_assign_ranks("=== ゲーム終了 ===")
            return

        for _ in range(n):
            next_pid = (self.current_pid + 1) % n
            if next_pid == 0:
                if not self.acted_in_round:
                    if not self.is_deck_completely_empty():
                        new_c = self.draw_card()
                        if new_c:
                            r_idx = random.randrange(len(self.field))
                            old_c = self.field[r_idx]
                            self.field[r_idx] = new_c
                            self.add_log(f"⏭️ 全員スキップ: 場の {old_c} に {new_c} を重ねました！")
                    if self.is_deck_completely_empty() and endless:
                        self.finish_game_and_assign_ranks("📭 山札（リフル上限含む）がなくなったためゲーム終了！")
                        return
                    elif self.is_deck_completely_empty() and not self.acted_in_round:
                        self.finish_game_and_assign_ranks("手詰まりのためゲーム終了！")
                        return
                self.turn += 1
                self.acted_in_round = False

            self.current_pid = next_pid
            if not self.players[self.current_pid]["finished"]:
                break

    def execute_ai_single_step(self):
        pid = self.current_pid
        p = self.players[pid]
        hand = p["hand"]
        endless = bool(self.settings.get("endlessMode", False))
        start_hand = int(self.settings.get("startHandSize", 5))
        end_limit = start_hand + int(self.settings.get("extraCard", 3))

        best_move = None
        best_score = -9999.0

        for h in hand:
            for f_idx, fc in enumerate(self.field):
                valid = self.find_common_anime(h, fc, pid)
                for anime in valid:
                    rem = len(hand) - 1
                    score = 0.0
                    if not endless:
                        if rem == 0:
                            score += 1000
                        score += (start_hand - rem) * 50
                    else:
                        score += (end_limit - rem) * 30
                        if rem == 1:
                            score += 20
                    co_cnt = sum(
                        1 for va in self.voice_actors if va != h and self.find_common_anime(h, va, pid)
                    )
                    score += (len(self.voice_actors) - co_cnt) * 0.1
                    if score > best_score:
                        best_score = score
                        best_move = (h, f_idx, anime)

        if best_move:
            h, f_idx, anime = best_move
            target = self.field[f_idx]
            r1 = get_first_character(anime, h)
            r2 = get_first_character(anime, target)
            self.add_log(f"{p['name']}: {h}（{r1}）を {target}（{r2}）に重ね『{anime}』を宣言")
            self.field[f_idx] = h
            hand.remove(h)
            self.used_anime.add(anime)
            p["usedAnime"].append(anime)
            self.acted_in_round = True
        else:
            if hand:
                discard_scores = []
                for idx_c, c in enumerate(hand):
                    val = sum(len(self.find_common_anime(c, fc, pid)) * 3 for fc in self.field)
                    val += sum(len(self.find_common_anime(c, o, pid)) for o in hand if o != c)
                    if endless and len(hand) >= end_limit - 1 and val <= 5:
                        val = -100
                    discard_scores.append((val, idx_c))
                discard_scores.sort()
                hand.pop(discard_scores[0][1])
                drawn = self.draw_card()
                if drawn:
                    hand.append(drawn)
            self.add_log(f"{p['name']}: パスして手札を1枚交換")

        self.check_victory_and_advance()


rooms: Dict[str, GameRoom] = {}


def gen_room_id() -> str:
    while True:
        rid = "".join(random.choices(string.digits, k=5))
        if rid not in rooms:
            return rid


async def run_ai_turns_if_needed(room: GameRoom):
    while room.phase == "PLAYING" and room.players[room.current_pid]["isAI"]:
        await asyncio.sleep(0.8)
        room.execute_ai_single_step()
        await room.broadcast()


@app.get("/")
async def index():
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/meta")
async def get_meta():
    return {"totalActors": len(all_actors)}


@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket):
    await ws.accept()
    client_id = ""
    current_room: Optional[GameRoom] = None

    try:
        while True:
            raw = await ws.receive_text()
            data = json.loads(raw)
            action = data.get("action")

            if action == "PING":
                await ws.send_text('{"type":"PONG"}')
                continue

            if action == "REGISTER_USER":
                client_id = data["clientId"]
                name = data.get("playerName", "").strip() or f"プレイヤー_{client_id[-4:]}"
                avatar = data.get("avatar", "")
                is_first_login = bool(data.get("isFirstLogin", False))

                existing = connected_users.get(client_id)
                room_id = existing.get("room_id") if existing else None
                connected_users[client_id] = {
                    "id": client_id,
                    "name": name,
                    "avatar": avatar,
                    "ws": ws,
                    "room_id": room_id,
                }
                if room_id and room_id in rooms:
                    r = rooms[room_id]
                    for h in r.lobby_humans:
                        if h["id"] == client_id:
                            h["name"] = name
                            h["avatar"] = avatar
                    if r.phase == "LOBBY":
                        await r.broadcast()
                await broadcast_online_users()
                if is_first_login:
                    await broadcast_user_login_toast(client_id, name, avatar)
                continue

            if action == "UPDATE_PROFILE":
                client_id = data["clientId"]
                name = data.get("playerName", "").strip() or f"プレイヤー_{client_id[-4:]}"
                avatar = data.get("avatar", "")
                if client_id in connected_users:
                    connected_users[client_id]["name"] = name
                    connected_users[client_id]["avatar"] = avatar
                if current_room:
                    for h in current_room.lobby_humans:
                        if h["id"] == client_id:
                            h["name"] = name
                            h["avatar"] = avatar
                    for p in current_room.players:
                        if p["id"] == client_id:
                            p["name"] = name
                            p["avatar"] = avatar
                    await current_room.broadcast()
                else:
                    await broadcast_online_users()
                continue

            if action == "CREATE_ROOM":
                client_id = data["clientId"]
                name = data.get("playerName", "").strip() or f"プレイヤー_{client_id[-4:]}"
                avatar = data.get("avatar", "")
                settings = data["settings"]
                rid = gen_room_id()
                room = GameRoom(rid, client_id, settings)
                room.connections[client_id] = ws
                room.lobby_humans.append({"id": client_id, "name": name, "avatar": avatar})
                rooms[rid] = room
                current_room = room
                connected_users[client_id] = {
                    "id": client_id,
                    "name": name,
                    "avatar": avatar,
                    "ws": ws,
                    "room_id": rid,
                }
                await room.broadcast()

            elif action == "JOIN_ROOM":
                client_id = data["clientId"]
                rid = data.get("roomId", "").strip()
                name = data.get("playerName", "").strip() or f"プレイヤー_{client_id[-4:]}"
                avatar = data.get("avatar", "")
                room = rooms.get(rid)
                if not room:
                    await ws.send_text(json.dumps({"type": "ERROR", "message": "ルームが見つかりません。"}, ensure_ascii=False))
                    continue

                is_existing = any(h["id"] == client_id for h in room.lobby_humans) or any(
                    p["id"] == client_id for p in room.players
                )
                if room.phase != "LOBBY" and not is_existing:
                    await ws.send_text(json.dumps({"type": "ERROR", "message": "すでにゲームが開始されています。"}, ensure_ascii=False))
                    continue

                room.connections[client_id] = ws
                if not any(h["id"] == client_id for h in room.lobby_humans):
                    room.lobby_humans.append({"id": client_id, "name": name, "avatar": avatar})
                else:
                    for h in room.lobby_humans:
                        if h["id"] == client_id:
                            h["name"] = name
                            h["avatar"] = avatar
                current_room = room
                connected_users[client_id] = {
                    "id": client_id,
                    "name": name,
                    "avatar": avatar,
                    "ws": ws,
                    "room_id": rid,
                }
                await room.broadcast()

            elif action == "LEAVE_ROOM":
                client_id = data.get("clientId", client_id)
                if current_room:
                    if client_id in current_room.connections:
                        del current_room.connections[client_id]
                    current_room.lobby_humans = [
                        h for h in current_room.lobby_humans if h["id"] != client_id
                    ]

                    was_current_turn = (
                        current_room.phase == "PLAYING"
                        and current_room.players[current_room.current_pid]["id"] == client_id
                    )

                    for p in current_room.players:
                        if p["id"] == client_id:
                            p["leftRoom"] = True
                            if not p["finished"]:
                                p["finished"] = True
                                p["rank"] = sum(1 for x in current_room.players if x["finished"])
                            current_room.add_log(f"🚪 {p['name']} がルームを退出しました。")

                    if current_room.host_id == client_id and current_room.lobby_humans:
                        current_room.host_id = current_room.lobby_humans[0]["id"]

                    if client_id in connected_users:
                        connected_users[client_id]["room_id"] = None

                    room_to_notify = current_room
                    current_room = None
                    await ws.send_text('{"type":"LEFT_ROOM_OK"}')

                    if room_to_notify.phase == "PLAYING":
                        active = [x for x in room_to_notify.players if not x["finished"]]
                        humans_active = [x for x in active if not x["isAI"]]
                        n = len(room_to_notify.players)
                        if (n == 1 and len(active) == 0) or (n > 1 and (len(active) <= 1 or len(humans_active) == 0)):
                            room_to_notify.finish_game_and_assign_ranks("=== プレイヤー退出によりゲーム終了 ===")
                        elif was_current_turn:
                            room_to_notify.check_victory_and_advance()
                            asyncio.create_task(run_ai_turns_if_needed(room_to_notify))

                    await room_to_notify.broadcast()
                continue

            elif action == "RECONNECT":
                client_id = data.get("clientId", "")
                rid = data.get("roomId", "")
                name = data.get("playerName", "").strip()
                avatar = data.get("avatar", "")
                if client_id in connected_users:
                    connected_users[client_id]["ws"] = ws
                    if name:
                        connected_users[client_id]["name"] = name
                    if avatar:
                        connected_users[client_id]["avatar"] = avatar

                room = rooms.get(rid)
                if room:
                    is_member = any(h["id"] == client_id for h in room.lobby_humans) or any(
                        p["id"] == client_id and not p.get("leftRoom") for p in room.players
                    )
                    if is_member:
                        room.connections[client_id] = ws
                        current_room = room
                        if client_id in connected_users:
                            connected_users[client_id]["room_id"] = rid
                        await ws.send_text(json.dumps(room.get_state_for_client(client_id), ensure_ascii=False))
                        await broadcast_online_users()
                        continue
                await ws.send_text('{"type":"SESSION_EXPIRED"}')

            elif action in ("START_GAME", "RESTART_GAME") and current_room:
                if client_id == current_room.host_id:
                    current_room.start_game()
                    await current_room.broadcast()
                    asyncio.create_task(run_ai_turns_if_needed(current_room))

            elif action == "CHECK_ANIME" and current_room:
                pid = next(
                    (i for i, p in enumerate(current_room.players) if p["id"] == client_id), -1
                )
                hand_card = data["handCard"]
                field_card = data["fieldCard"]
                animes = current_room.find_common_anime(hand_card, field_card, pid)
                options = [
                    {
                        "anime": a,
                        "roleHand": get_first_character(a, hand_card),
                        "roleField": get_first_character(a, field_card),
                    }
                    for a in animes
                ]
                await ws.send_text(json.dumps({
                    "type": "ANIME_OPTIONS",
                    "handCard": hand_card,
                    "fieldCard": field_card,
                    "options": options,
                }, ensure_ascii=False))

            elif action == "PLAY_CARD" and current_room and current_room.phase == "PLAYING":
                pid = current_room.current_pid
                p = current_room.players[pid]
                if p["id"] != client_id:
                    continue
                hand_card = data["handCard"]
                field_idx = int(data["fieldIdx"])
                anime = data["anime"]
                force_play = bool(data.get("forcePlay", False))
                target = current_room.field[field_idx]

                if hand_card in p["hand"] and (
                    force_play or anime in current_room.find_common_anime(hand_card, target, pid)
                ):
                    if force_play:
                        display_anime = "未登録作品（強制出し）"
                        current_room.add_log(f"{p['name']}: {hand_card} を {target} に重ね『{display_anime}』を宣言")
                        p["usedAnime"].append(f"{hand_card}×{target}(未登録)")
                    else:
                        r1 = get_first_character(anime, hand_card)
                        r2 = get_first_character(anime, target)
                        current_room.add_log(f"{p['name']}: {hand_card}（{r1}）を {target}（{r2}）に重ね『{anime}』を宣言")
                        current_room.used_anime.add(anime)
                        p["usedAnime"].append(anime)

                    current_room.field[field_idx] = hand_card
                    p["hand"].remove(hand_card)
                    current_room.acted_in_round = True
                    current_room.check_victory_and_advance()
                    await current_room.broadcast()
                    asyncio.create_task(run_ai_turns_if_needed(current_room))

            elif action == "SKIP_TURN" and current_room and current_room.phase == "PLAYING":
                pid = current_room.current_pid
                p = current_room.players[pid]
                if p["id"] != client_id:
                    continue
                mode = data.get("mode", "EXCHANGE")
                if mode == "KEEP" and len(p["hand"]) == 1:
                    current_room.add_log(f"{p['name']}: 手札1枚のため保持してスキップ")
                else:
                    d_idx = int(data.get("discardIdx", 0))
                    if 0 <= d_idx < len(p["hand"]):
                        discarded = p["hand"].pop(d_idx)
                        drawn = current_room.draw_card()
                        if drawn:
                            p["hand"].append(drawn)
                        current_room.add_log(f"{p['name']}: {discarded} を捨てて1枚引きました")
                current_room.check_victory_and_advance()
                await current_room.broadcast()
                asyncio.create_task(run_ai_turns_if_needed(current_room))

    except WebSocketDisconnect:
        if current_room and current_room.connections.get(client_id) is ws:
            del current_room.connections[client_id]
        if client_id in connected_users and connected_users[client_id].get("ws") is ws:
            del connected_users[client_id]
            await broadcast_online_users()


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="0.0.0.0", port=8000, reload=False)