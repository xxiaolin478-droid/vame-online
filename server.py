import asyncio
import datetime
import gc
import json
import random
import string
import sys
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import openpyxl
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI()

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
COAPPEAR_FILE = BASE_DIR / "共演表.xlsx"
CHARACTER_FILE = BASE_DIR / "キャラ表2.xlsx"

# --------- 起動時にExcelを省メモリモードで読み込み＆整数ID化 ---------
print("Excelファイルを省メモリモードで読み込んでいます...")

all_actors: List[str] = []
actor_to_id: Dict[str, int] = {}

all_animes: List[str] = []
anime_to_id: Dict[str, int] = {}


def get_or_add_anime_id(title: str) -> int:
    t = sys.intern(title)
    aid = anime_to_id.get(t)
    if aid is None:
        aid = len(all_animes)
        all_animes.append(t)
        anime_to_id[t] = aid
    return aid


# (anime_id, actor_id) -> 最初のキャラ名（文字列インターン済）
first_role_map: Dict[Tuple[int, int], str] = {}
# actor_id -> tuple of anime_ids
actor_anime_ids_map: Dict[int, List[int]] = {}

wb_char = openpyxl.load_workbook(CHARACTER_FILE, read_only=True, data_only=True)
ws_char = wb_char.active
char_rows = ws_char.iter_rows(values_only=True)

header_row = next(char_rows, None)
col_anime_ids: List[int] = []
if header_row:
    for c in header_row[1:]:
        title = str(c).strip() if c is not None else ""
        col_anime_ids.append(get_or_add_anime_id(title) if title else -1)

# 2行目（放送年行など）をスキップ
next(char_rows, None)

for row in char_rows:
    if not row or row[0] is None:
        continue
    actor_raw = str(row[0]).strip()
    if not actor_raw or actor_raw in ("nan", "NaN", "None"):
        continue
    actor = sys.intern(actor_raw)
    if actor in actor_to_id:
        actor_id = actor_to_id[actor]
    else:
        actor_id = len(all_actors)
        all_actors.append(actor)
        actor_to_id[actor] = actor_id
        actor_anime_ids_map[actor_id] = []

    for c_idx, cell in enumerate(row[1:]):
        if cell is None or c_idx >= len(col_anime_ids):
            continue
        aid = col_anime_ids[c_idx]
        if aid < 0:
            continue
        cell_str = str(cell).strip()
        if not cell_str or cell_str in ("nan", "NaN"):
            continue
        roles = [r.strip() for r in cell_str.split("`") if r.strip()]
        if roles:
            key = (aid, actor_id)
            if key not in first_role_map:
                first_role_map[key] = sys.intern(roles[0])
            actor_anime_ids_map[actor_id].append(aid)

wb_char.close()
del col_anime_ids
gc.collect()

# 無向ペア (min_actor_id, max_actor_id) -> Tuple[anime_id, ...]
coappear_pair_map: Dict[Tuple[int, int], Tuple[int, ...]] = {}
_temp_pair_sets: Dict[Tuple[int, int], Set[int]] = {}
_temp_coactors: Dict[Tuple[int, int], Set[int]] = {}

wb_co = openpyxl.load_workbook(COAPPEAR_FILE, read_only=True, data_only=True)
ws_co = wb_co.active
co_rows = ws_co.iter_rows(values_only=True)

co_header = next(co_rows, None)
col_actor_ids: List[int] = []
if co_header:
    for c in co_header[1:]:
        name = str(c).strip() if c is not None else ""
        col_actor_ids.append(actor_to_id.get(name, -1))

for row in co_rows:
    if not row or row[0] is None:
        continue
    a1_name = str(row[0]).strip()
    id1 = actor_to_id.get(a1_name, -1)
    if id1 < 0:
        continue

    for j, cell in enumerate(row[1:]):
        if cell is None or j >= len(col_actor_ids):
            continue
        id2 = col_actor_ids[j]
        if id2 < 0 or id1 == id2:
            continue
        cell_str = str(cell).strip()
        if cell_str in ("-", "", "nan", "NaN"):
            continue

        animes = [x.strip() for x in cell_str.split("`") if x.strip()]
        if not animes:
            continue

        pair_key = (id1, id2) if id1 < id2 else (id2, id1)
        a_set = _temp_pair_sets.setdefault(pair_key, set())
        for anime_title in animes:
            aid = get_or_add_anime_id(anime_title)
            a_set.add(aid)
            _temp_coactors.setdefault((aid, id1), set()).add(id2)
            _temp_coactors.setdefault((aid, id2), set()).add(id1)

wb_co.close()
del col_actor_ids

# ソート済みタプルに変換して省メモリ化
for pair_key, a_set in _temp_pair_sets.items():
    coappear_pair_map[pair_key] = tuple(sorted(a_set, key=lambda x: all_animes[x]))

del _temp_pair_sets

coactor_count_map: Dict[Tuple[int, int], int] = {}
anime_max_coactor_map: Dict[int, int] = {}

for (aid, act_id), co_set in _temp_coactors.items():
    cnt = len(co_set)
    coactor_count_map[(aid, act_id)] = cnt
    if cnt > anime_max_coactor_map.get(aid, 0):
        anime_max_coactor_map[aid] = cnt

del _temp_coactors

for aid in anime_max_coactor_map:
    if anime_max_coactor_map[aid] <= 0:
        anime_max_coactor_map[aid] = 1

# /api/meta 用キャッシュ（起動時に1度だけ生成）
META_RESPONSE_JSON = json.dumps(
    {
        "totalActors": len(all_actors),
        "actors": all_actors,
        "actorAnimes": {
            all_actors[act_id]: [all_animes[aid] for aid in aids]
            for act_id, aids in actor_anime_ids_map.items()
        },
    },
    ensure_ascii=False,
)
del actor_anime_ids_map
gc.collect()

print(f"読み込み完了！ (全声優数: {len(all_actors)}名 / 全作品数: {len(all_animes)}作)")

UNKNOWN_ROLE = sys.intern("？")


def get_first_character(anime_id: int, actor_id: int) -> str:
    return first_role_map.get((anime_id, actor_id), UNKNOWN_ROLE)


# --------- グローバル接続ユーザー＆全体チャット管理 ---------
connected_users: Dict[str, dict] = {}
global_chat_messages: List[dict] = []


def add_global_chat(client_id: str, name: str, avatar: str, text: str):
    jst = datetime.timezone(datetime.timedelta(hours=9))
    now_str = datetime.datetime.now(jst).strftime("%H:%M")
    global_chat_messages.append({
        "clientId": client_id,
        "name": name,
        "avatar": avatar,
        "text": text,
        "time": now_str,
    })
    if len(global_chat_messages) > 80:
        del global_chat_messages[0]


async def broadcast_global_chat():
    payload = json.dumps(
        {"type": "GLOBAL_CHAT_UPDATE", "chatMessages": global_chat_messages},
        ensure_ascii=False,
    )
    for u in list(connected_users.values()):
        ws = u.get("ws")
        if ws:
            try:
                await ws.send_text(payload)
            except Exception:
                pass


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


# --------- ルーム管理クラス（__slots__ で省メモリ化） ---------
class GameRoom:
    __slots__ = (
        "room_id",
        "host_id",
        "settings",
        "phase",
        "connections",
        "lobby_humans",
        "ai_configs",
        "match_count",
        "chat_messages",
        "players",
        "voice_actor_ids",
        "deck",
        "deck_refill_count",
        "max_refill",
        "field",
        "turn",
        "current_pid",
        "acted_in_round",
        "used_anime_ids",
        "logs",
        "ai_known_sets",
        "ai_task_running",
        "last_active",
    )

    def __init__(self, room_id: str, host_id: str, settings: dict):
        self.room_id = room_id
        self.host_id = host_id
        self.settings = settings
        self.phase = "LOBBY"  # LOBBY | PLAYING | RESULT
        self.connections: Dict[str, WebSocket] = {}
        self.lobby_humans: List[dict] = []
        self.ai_configs: List[dict] = settings.get("aiList", [])
        self.match_count = 0
        self.chat_messages: List[dict] = []

        self.players: List[dict] = []
        self.voice_actor_ids: List[int] = []
        self.deck: List[int] = []
        self.deck_refill_count = 0
        self.max_refill = 10
        self.field: List[int] = []
        self.turn = 1
        self.current_pid = 0
        self.acted_in_round = False
        self.used_anime_ids: Set[int] = set()
        self.logs: List[str] = []
        # pid -> Set[(min_actor_id, max_actor_id, anime_id)]
        self.ai_known_sets: Dict[int, Set[Tuple[int, int, int]]] = {}
        self.ai_task_running = False
        self.last_active = datetime.datetime.now(datetime.timezone.utc)

    def touch(self):
        self.last_active = datetime.datetime.now(datetime.timezone.utc)

    def add_chat(self, client_id: str, name: str, avatar: str, text: str):
        self.touch()
        jst = datetime.timezone(datetime.timedelta(hours=9))
        now_str = datetime.datetime.now(jst).strftime("%H:%M")
        self.chat_messages.append({
            "clientId": client_id,
            "name": name,
            "avatar": avatar,
            "text": text,
            "time": now_str,
        })
        if len(self.chat_messages) > 80:
            del self.chat_messages[0]

    def add_log(self, msg: str):
        self.logs.append(msg)
        if len(self.logs) > 60:
            del self.logs[0]

    def draw_card(self) -> Optional[int]:
        if not self.deck:
            if self.deck_refill_count < self.max_refill:
                self.deck = self.voice_actor_ids.copy()
                random.shuffle(self.deck)
                self.deck_refill_count += 1
                self.add_log(f"山札をリフル ({self.deck_refill_count}/{self.max_refill})")
            else:
                return None
        return self.deck.pop()

    def is_deck_completely_empty(self) -> bool:
        return len(self.deck) == 0 and self.deck_refill_count >= self.max_refill

    def find_common_anime_ids(self, id1: int, id2: int, pid: int) -> List[int]:
        if pid < 0 or pid >= len(self.players) or id1 == id2:
            return []
        pair_key = (id1, id2) if id1 < id2 else (id2, id1)
        all_ids = coappear_pair_map.get(pair_key, ())
        if not all_ids:
            return []

        player = self.players[pid]
        if not player["isAI"]:
            return list(all_ids)
        else:
            known = self.ai_known_sets.get(pid)
            if not known:
                return []
            return [aid for aid in all_ids if (pair_key[0], pair_key[1], aid) in known]

    def start_game(self):
        self.touch()
        self.match_count += 1
        selected_list = self.settings.get("selectedActors")
        if isinstance(selected_list, list) and len(selected_list) >= 4:
            self.voice_actor_ids = [actor_to_id[a] for a in selected_list if a in actor_to_id]
        else:
            limit = max(4, min(int(self.settings.get("actorLimit", 220)), len(all_actors)))
            self.voice_actor_ids = list(range(limit))

        if len(self.voice_actor_ids) < 4:
            self.voice_actor_ids = list(range(min(220, len(all_actors))))

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

        # AI知識構築（選択された声優ペアのみ対象にして高速＆省メモリ化）
        self.ai_known_sets = {}
        ai_pids = [pid for pid, p in enumerate(self.players) if p["isAI"]]
        if ai_pids:
            va_ids = self.voice_actor_ids
            n_va = len(va_ids)
            for pid in ai_pids:
                self.ai_known_sets[pid] = set()

            for i in range(n_va):
                id1 = va_ids[i]
                for j in range(i + 1, n_va):
                    id2 = va_ids[j]
                    pair_key = (id1, id2) if id1 < id2 else (id2, id1)
                    aids = coappear_pair_map.get(pair_key)
                    if not aids:
                        continue
                    for aid in aids:
                        c1 = coactor_count_map.get((aid, id1), 0)
                        c2 = coactor_count_map.get((aid, id2), 0)
                        mc = anime_max_coactor_map.get(aid, 1)
                        b1 = 0.5 + 0.5 * ((c1 / mc) ** 2)
                        b2 = 0.5 + 0.5 * ((c2 / mc) ** 2)
                        factor = (b1 + b2) / 2.0
                        triple = (pair_key[0], pair_key[1], aid)
                        for pid in ai_pids:
                            if random.random() * 100.0 < self.players[pid]["aiPct"] * factor:
                                self.ai_known_sets[pid].add(triple)

        self.deck = self.voice_actor_ids.copy()
        random.shuffle(self.deck)
        self.deck_refill_count = 0
        self.used_anime_ids = set()
        self.logs = []
        self.turn = 1
        self.current_pid = 0
        self.acted_in_round = False

        start_hand = max(1, min(15, int(self.settings.get("startHandSize", 5))))
        for p in self.players:
            p["hand"] = [c for _ in range(start_hand) if (c := self.draw_card()) is not None]

        self.field = [c for _ in range(3) if (c := self.draw_card()) is not None]
        self.phase = "PLAYING"
        self.add_log(f"第{self.match_count}戦 開始！（使用声優: {len(self.voice_actor_ids)}名 / 順番シャッフル済）")

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
                "hand": [all_actors[cid] for cid in p["hand"]] if i == my_pid else [],
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
            "settings": {
                "endlessMode": self.settings.get("endlessMode", False),
                "startHandSize": self.settings.get("startHandSize", 5),
                "extraCard": self.settings.get("extraCard", 3),
                "selectedCount": len(self.settings.get("selectedActors") or []),
            },
            "lobbyHumans": self.lobby_humans,
            "aiConfigs": self.ai_configs,
            "myPid": my_pid,
            "turn": self.turn,
            "currentPid": self.current_pid,
            "field": [all_actors[fid] for fid in self.field],
            "deckCount": len(self.deck),
            "refillCount": self.deck_refill_count,
            "maxRefill": self.max_refill,
            "players": public_players,
            "logs": self.logs,
            "chatMessages": self.chat_messages,
        }

    async def broadcast(self):
        self.touch()
        for cid, ws in list(self.connections.items()):
            try:
                await ws.send_text(json.dumps(self.get_state_for_client(cid), ensure_ascii=False))
            except Exception:
                pass
        await broadcast_online_users()

    async def broadcast_chat_only(self):
        self.touch()
        payload = json.dumps(
            {"type": "CHAT_UPDATE", "chatMessages": self.chat_messages},
            ensure_ascii=False,
        )
        for ws in list(self.connections.values()):
            try:
                await ws.send_text(payload)
            except Exception:
                pass

    def finish_game_and_assign_ranks(self, reason_log: str):
        active = [x for x in self.players if not x["finished"]]
        for rem in active:
            rem["finished"] = True
            rem["rank"] = sum(1 for x in self.players if x["finished"])
        self.phase = "RESULT"
        self.ai_known_sets.clear()  # 結果画面に移行したらAI知識メモリを即解放
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
            if drawn is not None:
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
                    if not self.is_deck_completely_empty() and self.field:
                        new_c = self.draw_card()
                        if new_c is not None:
                            r_idx = random.randrange(len(self.field))
                            old_c = self.field[r_idx]
                            self.field[r_idx] = new_c
                            self.add_log(
                                f"⏭️ 全員スキップ: 場の {all_actors[old_c]} に {all_actors[new_c]} を重ねました！"
                            )
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
        if not p["isAI"] or p["finished"]:
            return

        hand = p["hand"]
        endless = bool(self.settings.get("endlessMode", False))
        start_hand = int(self.settings.get("startHandSize", 5))
        end_limit = start_hand + int(self.settings.get("extraCard", 3))

        best_move = None
        best_score = -9999.0

        # 手札カードごとの共演可能数をキャッシュして高速化
        co_cnt_cache: Dict[int, int] = {}

        for h_idx, h in enumerate(hand):
            for f_idx, fc in enumerate(self.field):
                valid_aids = self.find_common_anime_ids(h, fc, pid)
                if not valid_aids:
                    continue
                if h not in co_cnt_cache:
                    co_cnt_cache[h] = sum(
                        1 for va in self.voice_actor_ids if va != h and self.find_common_anime_ids(h, va, pid)
                    )
                co_cnt = co_cnt_cache[h]

                for aid in valid_aids:
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
                    score += (len(self.voice_actor_ids) - co_cnt) * 0.1
                    if score > best_score:
                        best_score = score
                        best_move = (h_idx, h, f_idx, aid)

        if best_move:
            h_idx, h, f_idx, aid = best_move
            target = self.field[f_idx]
            anime_title = all_animes[aid]
            r1 = get_first_character(aid, h)
            r2 = get_first_character(aid, target)
            self.add_log(
                f"{p['name']}: {all_actors[h]}（{r1}）を {all_actors[target]}（{r2}）に重ね『{anime_title}』を宣言"
            )
            self.field[f_idx] = h
            hand.pop(h_idx)
            self.used_anime_ids.add(aid)
            p["usedAnime"].append(anime_title)
            self.acted_in_round = True
        else:
            if hand:
                discard_scores = []
                for idx_c, c in enumerate(hand):
                    val = sum(len(self.find_common_anime_ids(c, fc, pid)) * 3 for fc in self.field)
                    val += sum(len(self.find_common_anime_ids(c, o, pid)) for o in hand if o != c)
                    if endless and len(hand) >= end_limit - 1 and val <= 5:
                        val = -100
                    discard_scores.append((val, idx_c))
                discard_scores.sort()
                hand.pop(discard_scores[0][1])
                drawn = self.draw_card()
                if drawn is not None:
                    hand.append(drawn)
            self.add_log(f"{p['name']}: パスして手札を1枚交換")

        self.check_victory_and_advance()


rooms: Dict[str, GameRoom] = {}


def cleanup_stale_rooms():
    now = datetime.datetime.now(datetime.timezone.utc)
    stale_ids = []
    for rid, r in list(rooms.items()):
        # 誰も接続しておらず、10分以上経過したルームを削除
        if not r.connections and (now - r.last_active).total_seconds() > 600:
            stale_ids.append(rid)
    for rid in stale_ids:
        rooms.pop(rid, None)
    if stale_ids:
        gc.collect()


def gen_room_id() -> str:
    cleanup_stale_rooms()
    while True:
        rid = "".join(random.choices(string.digits, k=5))
        if rid not in rooms:
            return rid


async def run_ai_turns_if_needed(room: GameRoom):
    # 多重起動防止ロック
    if room.ai_task_running:
        return
    room.ai_task_running = True
    try:
        while (
            room.phase == "PLAYING"
            and 0 <= room.current_pid < len(room.players)
            and room.players[room.current_pid]["isAI"]
            and not room.players[room.current_pid]["finished"]
        ):
            await asyncio.sleep(0.8)
            # スリープ中にフェーズが変わった・人間が退出して終了した等の再チェック
            if (
                room.phase != "PLAYING"
                or not (0 <= room.current_pid < len(room.players))
                or not room.players[room.current_pid]["isAI"]
            ):
                break
            room.execute_ai_single_step()
            await room.broadcast()
    finally:
        room.ai_task_running = False


@app.get("/")
async def index():
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/meta")
async def get_meta():
    from fastapi.responses import Response
    return Response(content=META_RESPONSE_JSON, media_type="application/json")


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
                avatar = data.get("avatar", "")[:35000]
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
                await ws.send_text(
                    json.dumps(
                        {"type": "GLOBAL_CHAT_UPDATE", "chatMessages": global_chat_messages},
                        ensure_ascii=False,
                    )
                )
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
                avatar = data.get("avatar", "")[:35000]
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

            if action == "SEND_CHAT":
                text = str(data.get("text", "")).strip()[:200]
                if text:
                    u = connected_users.get(client_id, {})
                    name = u.get("name") or data.get("playerName", "プレイヤー")
                    avatar = u.get("avatar", "")
                    if current_room:
                        current_room.add_chat(client_id, name, avatar, text)
                        await current_room.broadcast_chat_only()
                    else:
                        add_global_chat(client_id, name, avatar, text)
                        await broadcast_global_chat()
                continue

            if action == "CREATE_ROOM":
                client_id = data["clientId"]
                name = data.get("playerName", "").strip() or f"プレイヤー_{client_id[-4:]}"
                avatar = data.get("avatar", "")[:35000]
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
                avatar = data.get("avatar", "")[:35000]
                room = rooms.get(rid)
                if not room:
                    await ws.send_text(
                        json.dumps({"type": "ERROR", "message": "ルームが見つかりません。"}, ensure_ascii=False)
                    )
                    continue

                is_existing = any(h["id"] == client_id for h in room.lobby_humans) or any(
                    p["id"] == client_id for p in room.players
                )
                if room.phase != "LOBBY" and not is_existing:
                    await ws.send_text(
                        json.dumps({"type": "ERROR", "message": "すでにゲームが開始されています。"}, ensure_ascii=False)
                    )
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
                        and 0 <= current_room.current_pid < len(current_room.players)
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

                    if not room_to_notify.lobby_humans and not room_to_notify.connections:
                        rooms.pop(room_to_notify.room_id, None)
                        gc.collect()
                    else:
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
                avatar = data.get("avatar", "")[:35000]
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
                        if room.phase == "PLAYING":
                            asyncio.create_task(run_ai_turns_if_needed(room))
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
                hand_card = data.get("handCard", "")
                field_card = data.get("fieldCard", "")
                id1 = actor_to_id.get(hand_card, -1)
                id2 = actor_to_id.get(field_card, -1)
                aids = current_room.find_common_anime_ids(id1, id2, pid) if (id1 >= 0 and id2 >= 0) else []
                options = [
                    {
                        "anime": all_animes[aid],
                        "roleHand": get_first_character(aid, id1),
                        "roleField": get_first_character(aid, id2),
                    }
                    for aid in aids
                ]
                await ws.send_text(
                    json.dumps(
                        {
                            "type": "ANIME_OPTIONS",
                            "handCard": hand_card,
                            "fieldCard": field_card,
                            "options": options,
                        },
                        ensure_ascii=False,
                    )
                )

            elif action == "PLAY_CARD" and current_room and current_room.phase == "PLAYING":
                pid = current_room.current_pid
                p = current_room.players[pid]
                if p["id"] != client_id:
                    continue
                hand_card = data.get("handCard", "")
                field_idx = int(data.get("fieldIdx", -1))
                anime_title = data.get("anime", "")
                force_play = bool(data.get("forcePlay", False))

                if not (0 <= field_idx < len(current_room.field)):
                    continue
                hand_actor_id = actor_to_id.get(hand_card, -1)
                if hand_actor_id < 0 or hand_actor_id not in p["hand"]:
                    continue

                target_actor_id = current_room.field[field_idx]
                target_name = all_actors[target_actor_id]
                aid = anime_to_id.get(anime_title, -1)
                valid_aids = current_room.find_common_anime_ids(hand_actor_id, target_actor_id, pid)

                if force_play or (aid >= 0 and aid in valid_aids):
                    if force_play:
                        display_anime = "未登録作品（強制出し）"
                        current_room.add_log(f"{p['name']}: {hand_card} を {target_name} に重ね『{display_anime}』を宣言")
                        p["usedAnime"].append(f"{hand_card}×{target_name}(未登録)")
                    else:
                        r1 = get_first_character(aid, hand_actor_id)
                        r2 = get_first_character(aid, target_actor_id)
                        current_room.add_log(
                            f"{p['name']}: {hand_card}（{r1}）を {target_name}（{r2}）に重ね『{anime_title}』を宣言"
                        )
                        current_room.used_anime_ids.add(aid)
                        p["usedAnime"].append(anime_title)

                    current_room.field[field_idx] = hand_actor_id
                    p["hand"].remove(hand_actor_id)
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
                        discarded_id = p["hand"].pop(d_idx)
                        drawn_id = current_room.draw_card()
                        if drawn_id is not None:
                            p["hand"].append(drawn_id)
                        current_room.add_log(f"{p['name']}: {all_actors[discarded_id]} を捨てて1枚引きました")
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