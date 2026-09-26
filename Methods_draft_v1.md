# Methods (draft v1)

> 说明：本稿按《统计分析计划_SAP_v1》与方案 V2.0 撰写，对应 STROBE 条目 4–12。方括号 [ ] 内为尚待确认或须在裁定、分母数据完成后填入的内容，投稿前逐一核实并删除方括号。

## Study design and setting

This was a retrospective, single-centre observational study reported in accordance with the STROBE statement. It was conducted at Wuhan Children's Hospital, Tongji Medical College, Huazhong University of Science and Technology [confirm the official English name of the institution], a tertiary paediatric referral centre in Wuhan, China. The study window was 1 June 2016 to 31 May 2026. The study was approved by the institutional Medical Ethics Committee (approval No. 2026R090-E01; protocol amendment approved [date]), which waived the requirement for informed consent because only existing inpatient records were used and all data were de-identified before analysis.

## Case identification

Admissions were retrieved from the hospital research data platform, which integrates the discharge abstract, admission and operative notes, discharge summaries, imaging, endoscopy, pathology and laboratory reports [describe the search strategy: diagnosis names and/or ICD-10 codes and the record fields searched]. The platform contained no admissions for July 2019 to March 2020, both in this extraction and in an independent extraction of foreign-body admissions over the same period, whereas adjacent months were populated; these nine months were treated as unobserved and excluded from all time-at-risk calculations [update once the medical records department has confirmed whether data for this period exist].

To assess the completeness of case finding, the medical records department provided monthly counts of admissions carrying perforation-related diagnosis codes in any diagnostic position:
- K31.8 subcodes for gastric or duodenal perforation and gastric rupture;
- perforation subcodes (.1, .2, .5, .6) of K25–K27;
- neonatal gastric perforation (P78.8 subcode) and congenital defect of the gastric muscular wall (Q40.204);
- gastric injury (S36.3) and duodenal injury (S36.4 subcodes).

These counts were compared month by month with the retrieved admissions, and months with excess counts were searched again [report the number of additional cases identified, if any].

## Participants

Children younger than 18 years were eligible if gastric and/or duodenal perforation was confirmed at operation or endoscopy, or diagnosed on imaging together with compatible clinical findings. Children were included irrespective of treatment, including those managed non-operatively and those whose families declined surgery.

Admissions were excluded when:
- the perforation was not located in the stomach or duodenum;
- there was no active perforation during the index admission (e.g. admission for ulcer, follow-up imaging or tube removal only);
- the perforation had been repaired at another hospital and the admission addressed its complications;
- the admission was a readmission for the same perforation (the first admission was retained);
- the records were insufficient to determine the aetiology after review of the original hospital information system entries.

The unit of analysis was the perforation episode.

## Aetiological adjudication

Two investigators [initials] independently classified every admission using a pre-specified sheet of operational definitions. Each investigator worked from a separate copy of the case summaries that omitted the automated keyword flags used during case retrieval, and neither had access to the other's classification. Operative or endoscopic findings took precedence over the clinical history, and discharge codes were used only as supporting information.

Six aetiological categories were defined:
- **Foreign body:** a foreign body confirmed at operation or endoscopy, with the perforation at the site of lodgement; subclassified as magnetic or non-magnetic, with the number of magnets retrieved.
- **Neonatal spontaneous perforation:** age ≤28 days, with a defect of the gastric muscular wall or spontaneous gastric rupture and no foreign body, ulcer, trauma or preceding procedure.
- **Peptic ulcer:** a chronic ulcer base at operation or on histology.
- **Trauma:** a definite history of external force, with contusion or laceration at operation.
- **Iatrogenic:** perforation after a diagnostic or therapeutic procedure at a concordant site.
- **Other or undetermined:** including perforation associated with congenital anomalies, necrotising enterocolitis, bezoar or systemic disease.

Inter-rater agreement was quantified with Cohen's κ and its 95% confidence interval (CI), using the large-sample standard error of Fleiss, Cohen and Everitt. Fields that depend on a preceding judgement, such as the foreign-body type, were assessed only in cases where both investigators agreed on that preceding judgement. Disagreements were resolved by a third investigator [initials] after review of the original records.

For the analysis, trauma, iatrogenic and other or undetermined causes were combined into a single "other" category. This was pre-specified because each was expected to include fewer than 15 cases.

## Variables

The following variables were obtained automatically from the extracted records:
- sex and age at admission;
- length of stay and postoperative length of stay;
- readmission for the same perforation within 30 days of discharge;
- the first white blood cell count and C-reactive protein concentration within 48 hours of admission.

The following were recorded during adjudication:
- site and number of perforations;
- treatment approach (open surgery, laparoscopy, conversion to open surgery, endoscopy or non-operative management);
- in-hospital outcome.

The following were abstracted by one investigator using predefined operational definitions and verified case by case by a second investigator [initials]:
- clinical presentation: duration of symptoms before admission, abdominal pain, vomiting, fever, abdominal distension, peritoneal signs, shock or sepsis at admission, and pneumoperitoneum on radiography or CT before treatment;
- in-hospital complications, with the highest grade by the Clavien–Dindo classification;
- unplanned reoperation and non-operative reintervention;
- time to first and to full enteral feeding, counted from the day of the first operation, or from admission in children managed non-operatively.

Progress notes were not part of the data extraction. Items not documented in the admission, operative or discharge records were therefore sought in the hospital information system, and items that remained undocumented were recorded as unknown.

Age was grouped as neonate (≤28 days), infant or toddler (29 days to <3 years), preschool (3 to <6 years) and school age (≥6 years). A composite adverse outcome was defined as in-hospital death or discharge after the family declined further treatment.

Admission time was analysed as a continuous variable (calendar year plus month midpoint). It was also grouped into three pre-specified periods defined by the COVID-19 pandemic: June 2016–December 2019, January 2020–December 2022 and January 2023–May 2026. After exclusion of the unobserved months, these periods contributed 37, 33 and 41 months of observation, respectively.

## Denominators

The medical records department provided monthly admissions to the wards that admit children with this condition, counted by admission date and discharge ward. These wards were the neonatal surgery ward, the two general surgery wards, and their successors after the wards were renamed in [month] 2024. These admissions served as the primary denominator. Alternative denominators were:
- admissions to these wards plus the neonatal medical and intensive care wards;
- emergency admissions to the surgical wards;
- all hospital admissions [retain only those actually provided].

Foreign-body admissions were defined by an ICD-10 T18 code in any diagnostic position. They were obtained from an institutional extraction beginning in July 2016, supplemented by [4] perforation admissions carrying a T18 code that the extraction had not captured. All rates are hospital-based relative frequencies per 10,000 surgical admissions (or per 1,000 foreign-body admissions) and do not represent population incidence.

## Statistical analysis

A statistical analysis plan was finalised before the adjudication results were unblinded [date]. Continuous variables are presented as median (interquartile range) and categorical variables as number (percentage), with missing values reported for each variable.

**Primary analysis.** The primary analysis tested the temporal trend in the proportion of perforations caused by a foreign body. It used the Cochran–Armitage test with continuous admission time as the score, with a two-sided P value obtained from 100,000 permutations. The effect size was the odds ratio per year from logistic regression.

**Secondary analyses** were exploratory:
- the distribution of aetiology across periods and across age groups, compared with the Fisher–Freeman–Halton exact test (Monte Carlo, 20,000 samples), with Cramér's V reported for age group;
- multinomial logistic regression with peptic ulcer as the reference category, giving the relative risk ratio per year for each aetiology;
- Poisson regression of monthly case counts with the logarithm of monthly surgical admissions as an offset, replaced by negative binomial regression when the Pearson dispersion statistic exceeded 1.5, giving incidence rate ratios per year;
- the same model applied to foreign-body–related and magnet-related perforations with foreign-body admissions as the offset, to assess whether any increase exceeded the change in foreign-body admissions.

Clinical characteristics were compared across aetiologies with the Kruskal–Wallis test or Fisher's exact test. Treatment and in-hospital outcomes were summarised descriptively, without between-group inference, because few events were expected.

**Sensitivity analyses.** The pre-specified sensitivity analyses were:
- repeating the primary analysis after excluding neonates, who cannot have foreign-body perforation, so that changes in neonatal referrals cannot drive the proportion;
- repeating the primary analysis in cases classified without arbitration;
- using the six-category classification;
- using the alternative denominators;
- restricting the count models to the 74 consecutive months from April 2020 onwards;
- adding any cases identified through the code-based completeness check.

Missing clinical data were handled by complete-case analysis without imputation. The primary analysis was the only confirmatory test, at a two-sided significance level of 0.05. Secondary and sensitivity analyses were not adjusted for multiplicity and are interpreted through effect sizes and 95% CIs. Analyses were performed in Python [version] with pandas, SciPy and statsmodels [versions]. The random seed was fixed at 20260926.
