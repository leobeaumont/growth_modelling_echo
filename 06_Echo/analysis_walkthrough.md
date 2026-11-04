# Echo project: walkthrough of the analysis

This document explains, step by step, what was done in `echo_growth_analysis.ipynb`, why it was done, and what was found. For each step it also says **how relevant the step is to the Echo case**, so you can decide what to keep in the report and the slides.

**Relevance scale used below**

| Label | Meaning |
|---|---|
| **Essential** | Needed to answer the board's question, or required by the project brief (cleaning log, growth model with backtest and 3 scenarios, A/B test plan). Do not remove. |
| **Useful** | Supports the answer or answers a guiding question with a real finding. Keep, but it can be shortened. |
| **Optional** | Correct and interesting, but not needed for the Echo story. Can be removed from the report and slides without any problem (it can stay in the notebook as an appendix). |

---

## The short answer

**Echo should not roll out aggressive push notifications.** The "+30 % sessions" is real, but sessions are app opens, not listening. Users who receive ~32 pushes a week instead of 4 open the app more for a few weeks, then leave: 13 points fewer of them are still active after 15 weeks. Revenue per user is equal after 16 weeks and **18 % lower over two years**. Rolled out to everyone, the model predicts **−16 % revenue and −23 % weekly active users after 12 months**.

Two other findings matter for the strategy:
- **Activation is the real lever.** Users who follow 3 or more shows in their first week retain much better, but only 38 % of users get there.
- **Paid social does not pay back.** A paid user costs €2.61 and brings €2.04 of ad revenue over two years.

---

## Part A. The 5 guiding questions applied to Echo

The data pack says these questions are "starting points, not a checklist". Here is how well each one fits Echo.

| Guiding question | Fit with Echo | Why |
|---|---|---|
| 1. Funnel and growth loop, where does value leak? | **Relevant** | The push notification is the *trigger* of Echo's content loop, so the loop explains why a bad trigger hurts. The funnel reveals the activation leak (3 shows in week 1). |
| 2. Segments, does the difference survive like with like? | **Partly relevant** | Needed for the A/B test (does any segment benefit from aggressive push? No). The channel comparison gives a strong side finding on paid social. The "cities" part does not apply: the Echo data has no city. |
| 3. Unit economics: CAC, retention, LTV, payback | **Relevant, with limits** | Needed to turn the retention loss into euros (LTV). But CAC exists only for paid social (the other channels have no spend), and payback by cohort is flat, so part of the question is thin. |
| 4. What changed during the 18 months, and what did it cause? | **Weakly relevant** | Almost nothing changed except the A/B test: growth is steady, CAC is flat, cohorts are identical. The useful part is showing the test's footprint at platform level. Keep it short. |
| 5. Which North Star Metric, and why? | **Very relevant** | The whole case is a wrong-metric story: the growth team optimised sessions. Choosing listening hours as the North Star *is* the lesson. |

---

## Part B. Step by step

### Step 0. Setup

**What.** Imports, paths, constants (€9 CPM → €0.009 per ad, extraction date 30 June 2026, test window 15 Jan to 15 Mar 2026), and a chart style. The colour palette was checked with a colour-blindness validator, and every chart has direct labels or a legend.

**Relevance: technical only.** Nothing to present.

---

### Step 1. Data exploration of the raw extract

**What.** Read the three files as raw as possible (user file as text, so pandas cannot hide problems), then profiled them: row counts, missing values, distinct values, date formats, distributions and extreme values.

**Why.** The data pack warns that the extract is raw. You cannot write a cleaning log without first finding the problems.

**Found.**
- Row counts match the data pack exactly (355,325 rows in total).
- `users.csv` has 124 duplicate rows, two date formats (ISO and `DD/MM/YYYY`), four spellings of the push variant, and missing age buckets.
- `weekly_activity.csv` has 1,663 duplicate rows and impossible listening values: −1, 10,080 (7 × 24 × 60, a full week non-stop), and 99,999 (a placeholder).
- `marketing_spend.csv` covers only paid social, and May 2025 is missing.
- The data was extracted on a Tuesday, so the week of 29 June 2026 only covers two days.

**Relevance: Essential** as a prerequisite. In the report it needs only one or two sentences, because the cleaning log carries the detail.

---

### Step 2. Data cleaning and cleaning log

**What.** Each rule is applied and logged with the number of rows it affects. Checks that found nothing are logged too, because they prove the problem was looked for. The log is saved in `outputs/cleaning_log.csv`, and the clean files are in `outputs/clean/`.

| File | Issue | Rule | Rows |
|---|---|---|---|
| users | Exact duplicate rows | Dropped | 124 |
| users | Day-first dates (`DD/MM/YYYY`) | Parsed as day-first. Verified: the first part is sometimes above 12, the second never is | 311 |
| users | Variant labels `Control`, `AGGRESSIVE_PUSH`, `aggressive` | Mapped to `control` / `aggressive_push` | 211 |
| users | Missing age bucket | Kept as `unknown`. Dropping these users would bias the channel and variant mix | 657 |
| users | Variant inconsistent with test window | Check only: every user who signed up in the window is in the test, none outside | 0 |
| weekly | Exact duplicate rows | Dropped | 1,663 |
| weekly | Listening minutes −1 / 10,080 / 99,999 | Set to missing, then imputed from ad impressions | 95 / 105 / 100 |
| weekly | Partial last week | Rows kept, excluded from weekly KPIs and retention | 7,326 |
| spend | Missing May 2025 | Linear interpolation (€1,115); cost per paid user gives €1,112 | 1 |
| spend | No spend for organic and cross-promo | Assumed €0, flagged as a caveat | – |
| all | Internal or test accounts | Searched for several fingerprints. None found | 0 |

**Why each choice.**
- **Imputing rather than dropping listening minutes.** The rest of the row (sessions, ads, pushes) is valid. Ads are served at a very stable rate of about 8.6 per listening hour, so minutes can be rebuilt from ads.
- **Excluding the partial week.** Counting it would make every user look less active in the last week and bias retention down.
- **Internal or test accounts.** The data pack says to expect them, so the notebook checks for several fingerprints. These are malformed IDs, users with no activity, activity before signup, bot-like constant usage and impossible intensity. None appear. Saying so explicitly is better than silently skipping the point.

**Relevance: Essential**, because the cleaning log is graded. One nuance: the imputation affects only 300 of 332,736 rows (0.09 %). Simply dropping those rows would not change any result, so if space is short, the imputation can be summarised in one line.

---

### Step 2.5. Analysis method (helpers)

**What and why.** Three method choices that everything else depends on.
1. **Retention.** A user is "active in week k of life" if they have an activity row k weeks after their signup week.
2. **Right-censoring.** At week k, only users who *can* be observed that far are counted. Without this, recent cohorts would look like churners just because their future is not in the data yet. This is the main methodological trap of the dataset.
3. **Value curve.** `v(k)` is the ad revenue in week k divided by the users observable at k. Summing `v(k)` gives the expected revenue per acquired user, which is the LTV. Beyond the observed horizon (about 75 weeks), the active share is extended with an exponential decay fitted on the last 26 weeks. This extension is only 11 to 16 % of a two-year LTV.

**Relevance: Essential.** Explain it in two or three sentences in the report. It is a likely cold question ("how did you handle users who signed up last month?").

---

### Step 3. Funnel and growth loop (guiding question 1)

#### 3.1 Funnel by channel

**What.** An AARRR-style funnel on all standard users observable for at least 26 weeks (12,148 users who signed up in 2025).

**Found.**

| Step | All users | Paid social |
|---|---|---|
| Followed at least 1 show in week 1 | 87.8 % | 81.8 % |
| Followed at least 3 shows in week 1 | 38.0 % | 25.4 % |
| Active in week 4 | 78.7 % | 71.0 % |
| Active in week 12 | 50.8 % | 37.4 % |
| Active in week 26 | 32.7 % | 19.2 % |

**Relevance: Useful.** It sets up the activation finding and the paid social finding. It is not needed for the push decision itself.

#### 3.2 The "aha moment"

**What.** Week-12 retention by the number of shows followed in week 1.

**Found.** It is a step, not a slope. Following 0, 1 or 2 shows gives 40 to 44 % week-12 retention. Following 3 or more gives about 64 %. So **activation means following at least 3 shows in week 1**.

**Relevance: Essential.** It powers the optimistic scenario, the next A/B test, and the recommendation to invest in onboarding rather than notifications.

#### 3.3 Revenue concentration

**Found.** The top 20 % of users bring 51 % of revenue. Users still active at week 26 (33 % of signups) bring 70 % of revenue. Non-activated users are 62 % of signups but bring only half of revenue.

**Relevance: Optional.** It is a nice supporting fact for "retention is where the money is", but the story works without it.

#### 3.4 The growth loop

**What.** A diagram of Echo's loop: follow shows → new episodes → push about the new episode → session → listening → ads → revenue, which is reinvested in paid social. Podcast cross-promotion acts as a weak acquisition loop.

**Why.** It explains the mechanism behind the A/B result. The push is the *trigger* of the loop. A relevant trigger feeds the loop; an irrelevant one ("Don't miss out!" five times a day) creates empty sessions and wears the trigger out. The K-factor from the course cannot be measured, because there is no invite data.

**Relevance: Useful.** It works well as one diagram slide that frames the whole case.

---

### Step 4. Segments, comparing like with like (guiding question 2)

**What.**
- Retention curves by channel.
- Retention tables by channel, platform, age and activation.
- Week-8 retention by channel *within* each level of shows followed.
- Direct standardisation: every channel given the same shows-followed mix.
- A linear probability regression: active in week 8 as a function of channel, activation, shows, platform and age.

**Why.** Paid social users follow fewer shows, so their lower retention could be a pure mix effect. The data pack explicitly asks whether differences survive like with like.

**Found.**
- **The paid social gap survives.** Standardising the shows mix only shrinks the week-8 gap with organic from −18 to −14 points. The regression gives −14.6 points (t = −16) for paid social, at equal activation, platform and age.
- Podcast cross-promo users are the best channel: +7 points at equal activation.
- **The age gap does not survive.** 18-24 year-olds look worse only because they follow fewer shows. At equal activation, age has no significant effect.
- Platform has no effect.
- Activation adds +15 points of week-8 retention on top of the channel effect.

**Relevance.**
- Like-with-like channel comparison: **Useful.** It answers guiding question 2 with a real result and backs the paid social recommendation.
- Regression: **Optional.** The standardisation table already proves the point and is easier to explain orally. Keep the regression as backup for questions.
- Age and platform tables: **Optional.** One sentence is enough ("no platform effect; the age effect disappears once activation is controlled").
- Note that the segment analysis that matters most for the *board's question* is the A/B heterogeneity check in step 7.4, not this one.

---

### Step 5. Unit economics (guiding question 3)

#### 5.1 Revenue per active user

**Found.** An active user listens about 81 minutes a week and hears 11.5 ads, which is worth **€0.10 per week (€0.45 per month)**. This is the same on every channel, so **LTV differences come almost entirely from retention**.

**Relevance: Essential.** It is what converts retention losses into euros.

#### 5.2 CAC

**Found.** The paid social cost per acquired user is **€2.61 in every month**, while spend grew from €990 to €1,630 a month. There is no sign of saturation yet. The blended cost over all signups is €1.12, but that number hides the problem.

**Relevance: Useful.** Only paid social has a cost, so this is short.

#### 5.3 LTV, LTV:CAC and payback by channel

**Found.**

| Channel | LTV 1 year | LTV 2 years | LTV : CAC |
|---|---|---|---|
| Organic store | €2.55 | €3.65 | no media cost |
| Paid social | €1.60 | €2.04 | **0.78, never pays back** |
| Podcast cross-promo | €2.93 | €4.26 | no media cost |

Paid social recovers only a third of its CAC in the first 12 weeks, and this is the same in every quarter.

**Relevance: Useful, but secondary to the board question.** This is the kind of finding the brief rewards ("the best projects find what is not asked"). Keep one slide and one paragraph, and do not let it compete with the main answer.

#### 5.4 Cohort views (12-week value by quarter, monthly cohort triangle)

**Found.** All cohorts look the same over 18 months: about 85 % active the month after signup, 64 % in month 2, 44 % in month 4 and 25 % after 9 months. Neither the product nor paid social quality changed.

**Relevance: Optional.** The conclusion ("cohorts are stable") is worth one sentence because it rules out other explanations. The triangle chart mirrors the course slide, but it does not move the argument.

---

### Step 6. What changed during the 18 months (guiding question 4)

**What.** Monthly signups by channel with growth rates, the channel mix, and six weekly platform KPIs. The KPIs are weekly active users, revenue, pushes per active user, sessions per active user, minutes per active user and minutes per session.

**Found.**
- Steady growth: organic +4 % a month, paid +3 % (the same as the budget), cross-promo +5 %.
- The mix shifts away from paid social (46 % → 40 % of signups) towards cross-promo (18 % → 22 %).
- **The A/B test is visible at platform level.** Pushes per active user rose from 4 to 8.8 and stayed high, because test users kept receiving aggressive pushes after the window. Sessions per active user rose about 5 %, while minutes per active user did not move and minutes per session fell.
- No tracking incident.

**Relevance: Useful, keep short.** The honest answer to guiding question 4 is "nothing changed except the test". The six-panel KPI chart is good evidence for the "sessions up, listening flat" argument. The signups-by-channel chart is **Optional**.

---

### Step 7. The A/B test, the core of the case

#### 7.1 Validity check

**What and why.** Before reading any result, the notebook checks that the test is valid.
- There is no sample ratio mismatch: 1,366 control vs 1,343 aggressive, p = 0.66.
- The arms are balanced on channel, platform, age and activation (all p > 0.09).
- The treatment was really delivered: 32 pushes a week vs 4.
- Every test user can be observed for at least 16 weeks (weeks 0 to 15), so the comparison over that horizon is not affected by censoring.

**Relevance: Essential.** A board would ask "can we trust this test?" first.

#### 7.2 Reproducing the claim, then looking beyond it

**Found.**

| Per user | During the test window | Over the first 16 weeks of life |
|---|---|---|
| Sessions | **+34 %** | +8 % |
| Listening minutes | +2 % (not significant) | **−7 %** |
| Ad revenue | **+17 %** | **0 %** (€1.14 vs €1.14) |
| Weeks active | – | **−10 %** |

- The claim is true and even understated: revenue was up too in the short term.
- That short-term revenue gain came from **more ads per listening hour** (9.1 vs 8.5), probably an extra ad at each app open. It did not come from more listening.
- Minutes per session fell from 23.4 to 20.4: the pushes create short, empty sessions.

**Relevance: Essential.** This is the "why the growth team was misled" slide.

#### 7.3 Retention

**Found.** Aggressive users are slightly more active in week 2 (+2 points) and equal in week 4. Then the gap opens and keeps widening: −10 points in week 8, −12 in week 12, and **−13 in week 15** (45.5 % vs 32.7 %, a 28 % relative loss, p < 10⁻¹¹).

**Relevance: Essential.** This is the main evidence.

#### 7.4 Heterogeneity: does any segment benefit?

**Found.** The effect on week-12 retention is negative in **every** segment: every channel, both platforms, every age group, and both activated and non-activated users. Paid social users are hit hardest (−15 points).

**Relevance: Essential.** It rules out the fallback "then let's roll it out only to segment X".

#### 7.5 Cost of a rollout in euros

**What.** The observed value curves up to week 15 are extended with the standard decay shape and each arm's revenue per active week.

**Found.**
- Two-year LTV is €3.07 for control and €2.51 for aggressive (**−18 %**).
- 28 % fewer users are still active at week 26.
- Per 10,000 new users, a rollout means about −920 active users at week 26 and −€5,600 of revenue.

**Relevance: Essential.** It translates the finding into the board's language. Mention also that this is a **dark pattern** (fear-of-missing-out copy, five pushes a day), which is exactly the case in quiz 3/3 of the course's acquisition chapter.

---

### Step 8. North Star Metric (guiding question 5)

**What.** Candidate metrics are compared (sessions, weekly active users, ad revenue, listening hours), with correlations to revenue and the evidence from the test.

**Found and chosen: weekly listening hours.**
- It is the core user value: people listening to shows they chose.
- It drives revenue. One listening hour is worth about €0.077, and listening has the highest correlation with revenue (0.993).
- It cannot be inflated by pushes. In normal times sessions also correlate with revenue (0.98), but the test shows they can be inflated without any listening. Even short-term revenue was inflated.
- The current value is about **9,900 hours a week**, which is about 7,340 weekly active users × 1.35 hours each.
- A metric tree breaks it into input metrics: new users × activation × retention × listening per user.

**Relevance: Essential.** It is the conceptual answer to the case: the team chose the wrong metric. Had listening hours been the primary metric, the test would have read "flat" from day one and "negative" by week 8.

---

### Step 9. Growth model, backtest and scenarios (required deliverable)

#### 9.1 Model structure

**What.** This is a cohort version of the course formula *U(t+1) = U(t) × (1 − churn) + new users + referrals*. A single churn rate does not fit Echo, because churn is very high in the first weeks and low afterwards. So each weekly signup cohort, in each of 6 segments (3 channels × activated or not), decays along its own measured curve. Weekly active users and revenue are the sums over cohorts.
- Inputs: paid users come from spend divided by CAC; organic and cross-promo signups follow their trends; the share of activated users per channel.
- Outputs: weekly active users, revenue, and listening hours.

**Relevance: Essential** (required by the brief).

#### 9.2 Backtest

**What.** Every parameter is estimated with data available on 28 December 2025 only. The model then predicts January to June 2026, which is compared with reality.
- Version A forecasts signups as well.
- Version B uses the actual signups, which isolates the retention part.

**Found.**

| | Average error, active users | Average error, revenue |
|---|---|---|
| A: forecast signups | 0.9 % | 1.3 % |
| B: actual signups | 2.1 % | 1.7 % |

- Version B **over-predicts more and more after the test starts**, reaching +2.8 % by June. The model calibrated on 2025 does not know that half of the test users churn faster. **The backtest error is the footprint of the A/B test**, which independently confirms the test result.
- Version A looks better only because two errors cancel out: signups were under-forecast by 3.8 %.

**Relevance: Essential** (required). In the report, present version B as the main result and version A in one line.

#### 9.3 Three scenarios, July 2026 to June 2027

All scenarios start from the real user base, calibrated on the last observed week (calibration factor 0.97).

| Scenario | Active users, June 2027 | Listening hours per week, June 2027 | 12-month revenue | vs base |
|---|---|---|---|---|
| Pessimistic: aggressive push for everyone | 10,587 | 14,052 | €47,178 | −16 % |
| Base: status quo | 13,781 | 18,385 | €56,178 | – |
| Optimistic: guided onboarding, +10 points of activation | 14,200 | 18,939 | €57,118 | +1.7 % |

- **Pessimistic.** New users follow the retention ratio measured in the test. Existing users lose an extra 2.1 % a week (measured in weeks 8 to 15 of the test) for 26 weeks, which is an assumption. They also hear more ads per hour, so listening hours are computed with that rate.
- **Optimistic.** The gain is modest in year one because it applies only to new users, while 12-month revenue mostly comes from users Echo already has. It compounds later.

**Relevance: Essential** (required).

#### 9.4 What-if table and the value of activation

**What.** One parameter is changed at a time, as on course slide 98.

**Found.**
- Aggressive push is the worst option under every assumption. Even if only new users are affected, active users drop 14 % by June 2027. If existing users erode for a year, revenue drops 18 %.
- Increasing the paid budget by 50 % adds €4.9k of revenue for €11.8k of extra spend.
- Stopping paid social cuts active users by 21 % but improves revenue minus spend by €13.7k over 12 months.
- Activating one more user is worth €1.75 to €2.11 over two years. **An activated paid user pays back** (€3.33 against €2.61); a non-activated one does not (€1.59). This means paid social can be fixed rather than cut.

**Relevance.**
- The aggressive-push sensitivity lines: **Essential**. They cover the main untested assumption about existing users.
- The activation table: **Useful**. It backs the recommendation.
- The paid social budget lines: **Useful but secondary**. Keep them if the paid social finding is in the report.

---

### Step 10. Next A/B test plan (required deliverable)

**What.** A pre-registered test of a **"pick 3 shows" guided onboarding**, chosen because activation is the lever the data points to. The notebook includes a sample size calculation.

**Plan.**
- **Design:** new users split 50/50, stratified by channel and platform. Notifications are identical in both arms, so the effect is not confounded.
- **Primary metric:** week-8 retention, currently 61 %.
- **Secondary metrics:** activation rate, listening hours per user, revenue per user.
- **Guardrails:** onboarding completion, week-1 retention, ads per hour, push opt-outs, store rating.
- **Size:** 4,069 users per arm to detect +3 points (about 23 weeks of enrolment at 350 signups a week), or 2,276 per arm for +4 points (about 13 weeks). The activation mechanism can be read after about 9 weeks.
- **Decision rule:** ship if the primary metric improves with p < 0.05 and no guardrail drops by more than 1 point. No peeking before the planned sample size, and a weekly check of the 50/50 split.
- The plan also lists what the aggressive push test should have done: use listening hours as the primary metric, include retention guardrails, and wait at least 8 weeks of follow-up before reading.

**Relevance: Essential** (required).

---

### Step 11. Summary

A one-table recap and four recommendations:
1. Stop the aggressive variant.
2. Make listening hours the North Star.
3. Invest in activation.
4. Fix paid social before scaling it.

**Relevance: Useful.** It is the backbone of the report's conclusion.

---

## Part C. What to keep where

**Steps you can drop from the report and slides without losing anything:** revenue concentration (3.3), the regression (4), the age and platform tables (4), the 12-week value by quarter and the cohort triangle (5.4), and the signups-by-channel chart (6). They can stay in the notebook as an appendix.

**Suggested 10-slide storyline for the 8-minute defense:**
1. The board's question and the answer (no).
2. Echo's loop: the push is the trigger.
3. Data and cleaning in one line, plus the method (retention with censoring).
4. The claim reproduced: sessions +34 %, but listening flat.
5. Retention curves: −13 points at week 15.
6. No segment benefits (the heterogeneity chart).
7. In euros: LTV −18 %, and the pessimistic scenario at −16 % revenue.
8. North Star: weekly listening hours, with the metric tree.
9. Model backtest and the three scenarios.
10. Recommendations and the next A/B test (guided onboarding), with paid social as a side finding.

**Likely cold questions and where the answer is:**

| Question | Where to find the answer |
|---|---|
| "How did you handle recent users?" | Step 2.5, right-censoring |
| "Revenue went up during the test, why say no?" | Step 7.2 |
| "Why not roll it out to young users only?" | Step 7.4 |
| "What about existing users?" | Step 9.4 sensitivity lines |
| "Is the model reliable?" | Step 9.2 backtest |

---

## Part D. Assumptions and limits to state openly

- **LTV is gross ad revenue.** There is no cost data (hosting, content rights, ad-tech fees). Real margins would make paid social look worse, not better.
- **Organic and cross-promo are assumed free.** In reality cross-promotion probably has some cost.
- **The effect of aggressive push on existing users was never tested.** The pessimistic scenario assumes the erosion measured on new users lasts 26 weeks. The what-if table shows the range from 0 to 52 weeks, and the conclusion holds in every case.
- **The long-term tail is extrapolated** beyond about 75 weeks of observation. This is 11 to 16 % of a two-year LTV.
- **The optimistic scenario assumes** that users moved into activation by onboarding behave like naturally activated users. This is exactly what the proposed A/B test would verify.
- **No internal or test accounts were found**, even though the data pack mentions them. The checks are documented in the cleaning log.
