def calculate_points(game_data, rules, players):
    base = int(rules.get("BasePoint", 1))
    multiplier = 2 ** game_data.get("announcements", 0)
    round_points = (base + game_data.get("special_points", 0)) * multiplier
    
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

if __name__ == "__main__":
    players = ["A", "B", "C", "D"]
    rules = {"BasePoint": 1, "SoloMultiplier": 3}
    
    # Test Normal Win Re
    res = calculate_points({
        "type": "Normal", "winner_team": "Re", "re_players": ["A", "B"], 
        "announcements": 1, "special_points": 2
    }, rules, players)
    print(f"Normal Re Win (Base 1, Ann 1 (x2), Spec 2 -> 6): {res} (Sum: {sum(res.values())})")
    
    # Test Solo Win
    res2 = calculate_points({
        "type": "Solo", "winner_team": "Soloist", "soloist": "A",
        "announcements": 0, "special_points": 0
    }, rules, players)
    print(f"Solo Win (Base 1 -> A gets 3, others -1): {res2} (Sum: {sum(res2.values())})")
    
    # Test Solo Loss
    res3 = calculate_points({
        "type": "Solo", "winner_team": "Others", "soloist": "A",
        "announcements": 0, "special_points": 1
    }, rules, players)
    print(f"Solo Loss (Base 1, Spec 1 -> Total 2 -> A gets -6, others 2): {res3} (Sum: {sum(res3.values())})")
