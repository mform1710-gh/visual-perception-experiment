# housing_simulation_opt.py
# Run with: python housing_simulation_opt.py
# Requires: openpyxl, pandas, matplotlib (for saving outputs). If not needed, you can skip saving.

import random
import math
import csv
from copy import deepcopy

# ------- PARAMETERS (edit these) -------
YEARS = 40
TERM = 20                 # contract term in years
PMT = 24_000_000         # annual payment per contract (KRW)
INIT = 300_000_000       # initial one-time deposit for condition2 contract (KRW)
HOUSE_PRICE = 800_000_000
CHURN_RATE = 0.10        # annual voluntary cancellation fraction
REFUND_RATIO = 0.70      # on cancellation, refunded fraction of cum_paid
RETURN_RATE = 0.03       # annual compound rate used to compute maturity payout (for 20-year refund calc)
P5_MIN = 0.03
P5_MAX = 0.07
INITIAL_CASH = 1_000_000_000  # starting cash; change as needed
MAX_B_PER_YEAR = 50      # upper bound for random search (candidate space)
MAX_SEARCH_ITERS = 2000  # adjust search budget
SEED = 12345
SAVE_OUTPUTS = True      # set False to skip writing files

random.seed(SEED)

# ------- Helper classes & functions -------
class Cohort:
    """Represents a group of contracts that started in the same year."""
    def __init__(self, start_year, count, is_condition2=False):
        self.start_year = start_year
        self.count = float(count)   # can become fractional after churn; treat as expected value
        self.is_condition2 = is_condition2
        # cumulative paid so far by an average contract in this cohort (exclude INIT for B-type)
        # We'll track cum_paid as total per-person amount paid (PMT per year added each year)
        self.cum_paid = 0.0
        if is_condition2:
            # if condition2 initial deposit was paid at start, include it in cum_paid
            self.cum_paid += INIT

    def receive_annual_payment(self):
        self.cum_paid += PMT

def maturity_payout_per_person(is_condition2, term=TERM, return_rate=RETURN_RATE):
    """
    Compute the maturity payment per person for a cohort that completes TERM years.
    For the PMT annuity we use compound accumulation at return_rate over TERM:
      PMT * ( (1+r)^TERM - 1 ) / r
    For condition2 initial deposit, we compound INIT for TERM years: INIT * (1+r)^TERM
    Returns per-person payout amount at maturity time.
    """
    r = return_rate
    if r == 0:
        ann = PMT * TERM
    else:
        ann = PMT * (((1 + r)**TERM - 1) / r)
    if is_condition2:
        return ann + INIT * ((1 + r)**TERM)
    else:
        return ann

def simulate(B_list, buy_list, p5, verbose=False):
    """
    Simulate cashflow:
    - B_list: list len YEARS, integer number of condition1 contracts started that year
    - buy_list: list len YEARS, 0/1 whether company tries to buy a house that year (only possible if year < 19 and enough cash)
    - p5: annual fund return (compound) used throughout
    Return: (feasible_bool, balances_list, final_balance, houses_bought_list)
    Note: houses bought add a condition2 occupant immediately (we assume when buying a house, company signs that many C contracts as desired).
    """
    balance = float(INITIAL_CASH)
    cohorts = []  # list of Cohort objects
    houses = 0
    balances = []
    houses_bought = [0]*YEARS

    # conservative reserve: we'll compute an annual reserve buffer proportional to expected flows.
    # For simplicity here we compute a small yearly reserve equal to a fraction of expected PMTs across term.
    # This is tunable; main constraint is balance >=0 each year.
    for year in range(YEARS):
        # new B cohorts start this year
        if B_list[year] > 0:
            cohorts.append(Cohort(year, B_list[year], is_condition2=False))

        # Attempt house purchase if buy signal and before year 19 (0-indexed: buy allowed for year indices 0..18)
        if year < 19 and buy_list[year] == 1:
            if balance >= HOUSE_PRICE:
                balance -= HOUSE_PRICE
                houses += 1
                houses_bought[year] = 1
                # when house is bought, assume immediately it's filled by a condition2 contract:
                cohorts.append(Cohort(year, 1, is_condition2=True))
                # that new condition2 paid INIT already, but we'll reflect its PMT when we process inflows below
            else:
                # cannot buy due to insufficient cash -> buy attempt fails (no purchase)
                houses_bought[year] = 0

        # Collect annual PMT inflows from all active cohorts
        inflow = 0.0
        for c in cohorts:
            if c.count <= 0:
                continue
            # each active contract pays PMT annually (we model expected value; churn reduces counts later)
            inflow += c.count * PMT
            # accumulate the payment to cohort.cum_paid
            c.receive_annual_payment()

        # Add inflow to balance
        balance += inflow

        # Apply investment return on balance (compound)
        balance *= (1 + p5)

        # Apply churn: each cohort loses CHURN_RATE fraction this year; refunds are REFUND_RATIO * cum_paid
        total_refund = 0.0
        for c in cohorts:
            if c.count <= 0:
                continue
            exits = c.count * CHURN_RATE
            if exits > 0:
                # refund per exiting person = REFUND_RATIO * c.cum_paid
                refund_per = REFUND_RATIO * c.cum_paid
                total_refund += exits * refund_per
                c.count -= exits  # reduce expected active counts

        balance -= total_refund

        # Handle maturities: cohorts that started TERM years ago pay out and become inactive
        maturity_payout = 0.0
        for c in cohorts:
            if c.count <= 0:
                continue
            if (year - c.start_year) + 1 >= TERM:
                per_person = maturity_payout_per_person(c.is_condition2, term=TERM, return_rate=RETURN_RATE)
                maturity_payout += c.count * per_person
                c.count = 0.0  # cohort completed

        balance -= maturity_payout

        balances.append(balance)

        # Feasibility: if at any point balance < 0 -> infeasible
        if balance < 0:
            if verbose:
                print(f"Infeasible at year {year+1}, balance={balance}")
            return False, balances, balance, houses_bought

    # At end (year=YEARS), we do NOT sell houses; they remain as off-balance-sheet assets for G40
    final_balance = balances[-1] if len(balances) > 0 else balance
    return True, balances, final_balance, houses_bought

# ------- Simple stochastic search for a good plan -------
def random_plan_search(iterations=500, verbose=False):
    best = None
    best_value = -1e30
    for it in range(iterations):
        # sample random B and buy lists (within reasonable ranges)
        B = [random.randint(0, MAX_B_PER_YEAR) for _ in range(YEARS)]
        # ensure buy signals only in years 0..18 (19 years)
        k = random.randint(0, 10)  # try up to 10 purchases in random plans (tunable)
        buy = [0]*YEARS
        years = random.sample(range(0, 19), k)
        for y in years:
            buy[y] = 1
        p5 = random.uniform(P5_MIN, P5_MAX)
        feasible, balances, final_bal, houses_bought = simulate(B, buy, p5)
        if feasible and final_bal > best_value:
            best_value = final_bal
            best = (B, buy, p5, balances, final_bal, houses_bought)
        if verbose and it % 100 == 0:
            print(f"iter {it}: best {best_value:.0f}")
    return best

# ------- Run search and report -------
if __name__ == "__main__":
    print("Starting stochastic search... (this may take time depending on iterations)")
    candidate = random_plan_search(iterations=1000, verbose=True)

    if candidate is None:
        print("No feasible plan found in the random search.")
    else:
        B_best, buy_best, p5_best, balances_best, final_best, houses_bought = candidate
        print("=== Best plan found ===")
        print(f"Final cash (G40): {final_best:,.0f} KRW")
        print(f"Best p5 (annual fund return used): {p5_best:.4f}")
        print(f"Total houses purchased (sum of houses_bought): {sum(houses_bought)}")
        print("Buy-years:", [i+1 for i,v in enumerate(buy_best) if v==1])
        # Optionally save outputs
        if SAVE_OUTPUTS:
            try:
                # save CSV of year-by-year plan
                with open("best_config.csv", "w", newline="") as f:
                    w = csv.writer(f)
                    w.writerow(["Year", "B_condition1", "PurchasedHouseThisYear", "BuySignal"])
                    for y in range(YEARS):
                        w.writerow([y+1, B_best[y], houses_bought[y], buy_best[y]])
                # save simple balances CSV
                with open("simulation_balance.csv", "w", newline="") as f:
                    w = csv.writer(f)
                    w.writerow(["Year", "Balance"])
                    for y,bal in enumerate(balances_best):
                        w.writerow([y+1, bal])
                print("Outputs saved: best_config.csv, simulation_balance.csv")
            except Exception as e:
                print("Could not save outputs:", e)
