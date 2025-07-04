import aiohttp
import asyncio

API_KEY = "d9e0c483-4196-4f8c-b40a-6cefe7ab07d5"
HEADERS = {"Authorization": f"Bearer {API_KEY}"}

async def fetch(session, url):
    async with session.get(url, headers=HEADERS) as response:
        return await response.json()

async def get_player_id(session, nickname):
    url = f"https://open.faceit.com/data/v4/players?nickname={nickname}"
    data = await fetch(session, url)
    return data.get("player_id")

async def get_faceit_stats(session, player_id):
    url = f"https://open.faceit.com/data/v4/players/{player_id}"
    data = await fetch(session, url)
    try:
        return data["games"]["cs2"]["faceit_elo"]
    except KeyError:
        return "-"

async def get_lifetime_stats(session, player_id):
    url = f"https://open.faceit.com/data/v4/players/{player_id}/stats/cs2"
    data = await fetch(session, url)
    stats = data.get("lifetime", {})
    return {
        "kd": stats.get("Average K/D Ratio", "-"),
        "winrate": stats.get("Win Rate %", "-"),
        "matches": stats.get("Matches", "-"),
        "longest_win_streak": stats.get("Longest Win Streak", "-"),
        "average_headshots": stats.get("Average Headshots %", "-"),
        "recent_results": stats.get("Recent Results", "-")
    }

async def get_last_30_stats(session, player_id, nickname):
    history_url = f"https://open.faceit.com/data/v4/players/{player_id}/history?game=cs2&limit=30"
    history = await fetch(session, history_url)
    matches = history.get("items", [])

    result = {
        "kills": 0, "assists": 0, "deaths": 0, "headshots": 0,
        "adr": 0, "k_d_ratio": 0, "k_r_ratio": 0, "wins": 0
    }

    if not matches:
        return result

    tasks = []
    for match in matches:
        match_id = match["match_id"]
        stats_url = f"https://open.faceit.com/data/v4/matches/{match_id}/stats"
        tasks.append(fetch(session, stats_url))

    stats_results = await asyncio.gather(*tasks)

    for stats_json in stats_results:
        try:
            round_data = stats_json["rounds"][0]
            for team in round_data["teams"]:
                for player in team["players"]:
                    if player["nickname"].lower() == nickname.lower():
                        result["kills"] += int(player["player_stats"]["Kills"])
                        result["assists"] += int(player["player_stats"]["Assists"])
                        result["deaths"] += int(player["player_stats"]["Deaths"])
                        result["k_d_ratio"] += float(player["player_stats"]["K/D Ratio"])
                        result["k_r_ratio"] += float(player["player_stats"]["K/R Ratio"])
                        result["headshots"] += int(player["player_stats"]["Headshots %"])
                        result["adr"] += float(player["player_stats"]["ADR"])
                        result["wins"] += player["player_stats"]["Result"] == "1"
        except Exception as e:
            print(f"[❌] Ошибка в матче: {e}")

    result["losses"] = 30 - result["wins"]
    result["win_rate"] = round(result["wins"] / 30 * 100)
    result["kills"] = round(result["kills"] / 30)
    result["assists"] = round(result["assists"] / 30)
    result["deaths"] = round(result["deaths"] / 30)
    result["k_d_ratio"] = round(result["k_d_ratio"] / 30, 2)
    result["k_r_ratio"] = round(result["k_r_ratio"] / 30, 2)
    result["headshots"] = round(result["headshots"] / 30)
    result["adr"] = round(result["adr"] / 30, 2)
    
    return result

async def get_full_stats(nickname):
    async with aiohttp.ClientSession() as session:
        player_id = await get_player_id(session, nickname)
        if not player_id:
            return None

        elo = await get_faceit_stats(session, player_id)
        lifetime_stats = await get_lifetime_stats(session, player_id)
        last30 = await get_last_30_stats(session, player_id, nickname)

        result = {"elo": elo, **lifetime_stats, **last30}
        return result
