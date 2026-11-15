import numpy as np
import pandas as pd
from pathlib import Path

SEED = 7
START = pd.Timestamp("2026-07-01")
END = pd.Timestamp("2026-09-30 23:59")
RELEASE_WEEK = (pd.Timestamp("2026-08-10"), pd.Timestamp("2026-08-16"))
NEW_AGENT_START = pd.Timestamp("2026-08-03")

OUT = Path(__file__).resolve().parents[1] / "data"
rng = np.random.default_rng(SEED)

AGENTS = [
    ("Anna K.", "Team A"), ("David R.", "Team A"), ("Maria S.", "Team A"), ("Tom H.", "Team A"),
    ("Elena P.", "Team A"), ("Chris W.", "Team B"), ("Nina B.", "Team B"), ("Oliver J.", "Team B"),
    ("Sara L.", "Team B"), ("Mark T.", "Team B"), ("Julia F.", "Team B"), ("Lukas M.", "Team A"),
]
# agent skill: multiplier on resolution time, csat shift, reopen prob
SKILL = {
    "Anna K.": (0.8, 0.30, 0.04), "David R.": (1.0, 0.0, 0.06), "Maria S.": (0.9, 0.15, 0.05),
    "Tom H.": (1.2, -0.20, 0.07), "Elena P.": (0.85, 0.20, 0.04), "Chris W.": (1.0, 0.05, 0.06),
    "Nina B.": (0.95, 0.10, 0.05), "Oliver J.": (1.1, -0.10, 0.07), "Sara L.": (0.9, 0.20, 0.05),
    "Mark T.": (1.05, 0.0, 0.06), "Julia F.": (1.0, 0.05, 0.06), "Lukas M.": (1.6, -0.55, 0.15),
}

CHANNELS = {"Chat": 0.50, "Email": 0.35, "Phone": 0.15}
# share, median resolution hours, high-priority share, escalation prob
CATEGORIES = {
    "Payments":        (0.27, 7.0, 0.35, 0.07),
    "Account access":  (0.22, 2.0, 0.25, 0.02),
    "Technical issue": (0.24, 16.0, 0.20, 0.06),
    "Product question": (0.17, 1.2, 0.05, 0.01),
    "Refund request":  (0.10, 30.0, 0.15, 0.09),
}
FIRST_RESPONSE_MEDIAN_MIN = {"Chat": 4, "Email": 50, "Phone": 2}
FIRST_RESPONSE_SIGMA = {"Chat": 1.1, "Email": 0.7, "Phone": 0.9}


def daily_volume(day):
    base = 58 + 0.12 * (day - START).days
    if day.day_of_week >= 5:
        base *= 0.85
    if RELEASE_WEEK[0] <= day <= RELEASE_WEEK[1]:
        base *= 1.55
    return rng.poisson(base)


def hour_probs():
    h = np.array([1, 0.7, 0.5, 0.4, 0.4, 0.6, 1.2, 2.5, 4, 5, 5.5, 5.5,
                  5, 5.2, 5.3, 5, 4.8, 4.5, 4.2, 4, 3.6, 3, 2.2, 1.5])
    return h / h.sum()


def pick_priority(cat):
    high = CATEGORIES[cat][2]
    return rng.choice(["High", "Medium", "Low"], p=[high, (1 - high) * 0.6, (1 - high) * 0.4])


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    cats = list(CATEGORIES)
    cat_p = np.array([CATEGORIES[c][0] for c in cats])
    rows = []
    tid = 100001

    for day in pd.date_range(START, END.normalize(), freq="D"):
        weekend = day.day_of_week >= 5
        release = RELEASE_WEEK[0] <= day <= RELEASE_WEEK[1]
        available = [a for a, _ in AGENTS if a != "Lukas M." or day >= NEW_AGENT_START]

        for _ in range(daily_volume(day)):
            created = day + pd.Timedelta(hours=int(rng.choice(24, p=hour_probs())), minutes=int(rng.integers(0, 60)))
            channel = rng.choice(list(CHANNELS), p=list(CHANNELS.values()))
            p = cat_p.copy()
            if release:
                p[cats.index("Technical issue")] *= 2.2
            cat = rng.choice(cats, p=p / p.sum())
            priority = pick_priority(cat)
            agent = rng.choice(available)
            speed, csat_shift, reopen_p = SKILL[agent]

            fr_mult = (2.6 if weekend else 1.0) * (1.8 if release else 1.0)
            if created.hour < 7 or created.hour >= 22:
                fr_mult *= 2.0
            first_resp = rng.lognormal(np.log(FIRST_RESPONSE_MEDIAN_MIN[channel] * fr_mult), FIRST_RESPONSE_SIGMA[channel])
            if priority == "High":
                first_resp *= 0.6

            med = CATEGORIES[cat][1] * speed * (1.5 if release and cat == "Technical issue" else 1.0)
            res_hours = max(first_resp / 60 + 0.05, rng.lognormal(np.log(med), 0.8))

            escalated = rng.random() < CATEGORIES[cat][3] * (1.6 if agent == "Lukas M." else 1.0)
            if escalated:
                res_hours *= rng.uniform(2, 4)

            resolved_at = created + pd.Timedelta(hours=res_hours)
            status = "Resolved"
            if resolved_at > END:
                status, resolved_at = "Open", pd.NaT

            reopened = "Yes" if status == "Resolved" and rng.random() < reopen_p * (1.4 if cat == "Technical issue" else 1.0) else "No"

            csat = None
            if status == "Resolved" and rng.random() < 0.36:
                score = 4.6 + csat_shift - 0.45 * np.log1p(res_hours / 6) - (0.8 if reopened == "Yes" else 0) - (0.3 if escalated else 0)
                csat = int(np.clip(round(rng.normal(score, 0.7)), 1, 5))

            rows.append({
                "ticket_id": f"TCK-{tid}",
                "created_at": created,
                "first_response_at": created + pd.Timedelta(minutes=first_resp),
                "resolved_at": resolved_at,
                "channel": channel,
                "category": cat,
                "priority": priority,
                "agent": agent,
                "status": status,
                "escalated": "Yes" if escalated else "No",
                "reopened": reopened,
                "csat": csat,
            })
            tid += 1

    df = pd.DataFrame(rows)

    # export problems seen in real helpdesk data
    dup = df.sample(n=14, random_state=1)
    df = pd.concat([df, dup]).sort_values("created_at").reset_index(drop=True)
    idx = df[df.status == "Resolved"].sample(n=9, random_state=2).index
    df.loc[idx, "resolved_at"] = df.loc[idx, "created_at"] - pd.to_timedelta(rng.integers(1, 48, len(idx)), unit="h")
    df.loc[df.sample(n=17, random_state=3).index, "agent"] = None
    csat_idx = df[df.csat.notna()].sample(n=6, random_state=4).index
    df.loc[csat_idx, "csat"] = rng.choice([7, 8, 9, 10], len(csat_idx))

    df["csat"] = df["csat"].astype("Int64")
    df.to_csv(OUT / "helpdesk_tickets_q3_2026.csv", index=False, date_format="%Y-%m-%d %H:%M")
    pd.DataFrame(AGENTS, columns=["agent", "team"]).assign(
        start_date=lambda d: np.where(d.agent == "Lukas M.", "2026-08-03", "2025-03-01")
    ).to_csv(OUT / "agents.csv", index=False)
    print(f"{len(df):,} tickets written")


if __name__ == "__main__":
    main()
