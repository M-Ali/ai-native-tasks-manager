"""Code the collected Instagram posts into capture.xlsx.

Codes are authored from looking at the downloaded creative (contact sheets in
contact/). Anything whose message could not be read at thumbnail size is coded
`unclear` rather than guessed - re-open the full-size file in media/ to finish it.

    uv run --with openpyxl python workspace/audit/code_rows.py
"""

from __future__ import annotations

import csv
from pathlib import Path

import openpyxl

AUDIT = Path(__file__).parent

# shortcode -> (product_line, claim1, claim2, register, proof, cta, audience, objection, note)
C = {
    # ---------------- State Life (client) ----------------
    "Db40vJeDIdE": ("brand_corporate","national_trust","other","reassurance","claim_statistics","call_centre","mass","yes","E-Kachehri: 745 received -> 744 resolved. Bilingual. Grievance-redress framing, not advertising."),
    "Db5AI2Joex3": ("brand_corporate","brand_corporate","other","pride","expert_authority","none","sme_corporate","no","CEO at IAP/PSOA conference: 'Insurance: the quiet driver of growth'."),
    "Db_cOMoDCos": ("brand_corporate","national_trust","other","pride","none","none","mass","no","'Har Khwab Rahe Azaad' - 79 Years of Freedom. Calendar."),
    "DbfQOFYn5_V": ("brand_corporate","brand_corporate","other","pride","none","none","sme_corporate","no","CEO with Federal Minister. Grip-and-grin."),
    "DbOTNzjIDXw": ("brand_corporate","brand_corporate","protection","reassurance","heritage_scale","none","mass","no","'Protecting Every Life Journey' - product-range chart. A list, not a proposition."),
    "Dbx2qijjH7a": ("child_education","protection","education_marriage","reassurance","heritage_scale","digital_direct","mass","no","CAMPAIGN: 'Education Planning Starts Years Before the Admission Form / PLAN AHEAD WITH STATE LIFE'. Footer claims 'Pakistan's Largest Life and Health Insurer'."),
    "DbygIFcEQfj": ("individual_life","protection","savings_return","aspiration","heritage_scale","digital_direct","mass","no","CAMPAIGN: 'The best opportunities are easier to seize when you're prepared / PLAN AHEAD'. Skydiving."),
    "DbzJUzCgQRC": ("health","protection","other","reassurance","heritage_scale","digital_direct","mass","no","CAMPAIGN: 'Healthcare planning starts before you ever need care / PLAN AHEAD'."),
    "DcByD4DDwQZ": ("brand_corporate","national_trust","other","pride","none","none","mass","no","Flag-raising event photo."),
    "Dcd2vnBE2WW": ("brand_corporate","faith_compliance","other","reassurance","none","none","mass","no","Eid Milad-un-Nabi greeting."),
    "DcdSRaYDJgx": ("brand_corporate","brand_corporate","other","pride","none","none","youth","no","Qimam Fellowship students at Principal Office. Institutional PR."),
    "DcK3NfajCKE": ("brand_corporate","brand_corporate","other","pride","none","none","sme_corporate","no","Divisional Head at infrastructure conference. Grip-and-grin."),

    # ---------------- Jubilee Life ----------------
    "DaKLEGSh1X5": ("brand_corporate","national_trust","protection","reassurance","claim_statistics","none","mass","yes","KEY: 'PKR 54 BILLION CLAIMS PAID IN 2025 / Behind every claim is a promise fulfilled.' Pinned."),
    "DcaZMaQiguj": ("health","price_affordability","convenience_digital","aspiration","none","digital_direct","mass","no","Jubilee Active: 40% discount, laboratory."),
    "DcbbO6GsuM_": ("health","convenience_digital","other","aspiration","testimonial","digital_direct","youth","no","Creator UGC, Jubilee Active."),
    "Dcbi4YpNL_0": ("health","convenience_digital","other","aspiration","testimonial","digital_direct","youth","no","UGC: 'I have been using Jubilee Active'."),
    "Dcc1kfVAQlg": ("health","price_affordability","convenience_digital","aspiration","none","digital_direct","mass","no","10% discount - labs, radiology, consultations."),
    "Dcc1kIiCaRT": ("brand_corporate","other","other","aspiration","none","none","mass","no","Dark interior teaser, message not readable at thumbnail size."),
    "Dcdf27ktk9U": ("health","convenience_digital","other","aspiration","testimonial","digital_direct","youth","no","'Food/Health' creator content."),
    "DcdGTrIiAY8": ("health","price_affordability","convenience_digital","aspiration","none","retail","youth","no","Jubilee Active x Cafe D Grill, 40% OFF, Saffron card."),
    "DcdnGakiGqB": ("brand_corporate","brand_corporate","other","aspiration","none","none","youth","no","Iqra University Job Fair. Recruitment."),
    "DcdXP-dJM0H": ("health","convenience_digital","other","aspiration","testimonial","digital_direct","youth","no","Creator UGC, fitness."),
    "DceHneNo1lK": ("health","convenience_digital","price_affordability","aspiration","none","digital_direct","youth","no","Jubilee Active app: steps, points, redeem."),
    "DciVEastwvH": ("health","convenience_digital","other","aspiration","testimonial","digital_direct","women","no","Creator UGC, Jubilee Active."),

    # ---------------- EFU Life ----------------
    "Dca_LdWEWTe": ("brand_corporate","convenience_digital","savings_return","aspiration","none","digital_direct","youth","no","Thrive: 'Learn, earn and thrive... jeetain 1000 EFU COINS'."),
    "DcazN3EASRP": ("health","convenience_digital","other","aspiration","none","digital_direct","youth","no","WIN: 'What is your one WIN for today?'"),
    "DcI7gFrAP1p": ("brand_corporate","convenience_digital","savings_return","aspiration","none","digital_direct","youth","no","Thrive sign-up, 2000 EFU COINS, 'no paper work'."),
    "DcI9D_fEq-g": ("health","other","other","aspiration","expert_authority","none","women","no","Falak Zehra, Wellness Centre KSBL. Expert talking head."),
    "DciavMIFF7k": ("brand_corporate","price_affordability","other","aspiration","none","retail","youth","no","PRIMUS x Big Bash burger, 25% OFF."),
    "DcK-QqLHBqA": ("brand_corporate","brand_corporate","other","pride","none","none","sme_corporate","no","Staff group photo."),
    "DcLWVP1DdZh": ("brand_corporate","price_affordability","convenience_digital","aspiration","none","digital_direct","youth","no","WIN x KraveMart, Rs.500 vouchers."),
    "DcN7TnYjVis": ("brand_corporate","brand_corporate","other","aspiration","none","none","sme_corporate","no","Nexus Pro management development programme."),
    "DcNp6fnjfJU": ("health","convenience_digital","other","reassurance","none","digital_direct","sme_corporate","no","EFU Life mHealth - 'supporting specialists through connected healthcare'."),
    "DcOZlsPGrnn": ("health","price_affordability","other","aspiration","none","retail","mass","no","PRIMUS: 'Your health matters', up to 30% off diagnostics."),
    "DcVip3mjbwg": ("health","price_affordability","other","aspiration","none","retail","women","no","PRIMUS: 'Confidence starts here', up to 20% off clinic."),
    "DVdIMC4DE8k": ("brand_corporate","national_trust","protection","reassurance","expert_authority","none","mass","no","CEO Mohammed Ali Ahmed: 'STANDING STRONG TOGETHER... safeguarding you'. Closest EFU gets to a trust claim."),

    # ---------------- Adamjee Life ----------------
    "Db_cNbHD-Xo": ("brand_corporate","national_trust","other","pride","none","none","mass","no","'The Home to Every Life' - map of Pakistan, heritage illustration."),
    "Db_kUIJjaX7": ("brand_corporate","brand_corporate","other","pride","none","none","sme_corporate","no","Nigehbaan event photos."),
    "DcAaAdSDE4I": ("brand_corporate","national_trust","other","pride","heritage_scale","none","mass","no","70 years + 14 August cake-cutting."),
    "DcajWyjje2Z": ("brand_corporate","convenience_digital","other","reassurance","none","digital_direct","mass","no","'Kabhi bhi Kahin bhi' - Quick Pay app, 'Branch visit bhool jao'."),
    "DcAxZVCjv1L": ("brand_corporate","national_trust","other","pride","none","none","mass","no","Independence Day office decor."),
    "DcEXyLqIIf9": ("brand_corporate","brand_corporate","other","pride","none","none","sme_corporate","no","Cake-cutting event, letterboxed video."),
    "DcQg_ZWDSdB": ("brand_corporate","convenience_digital","other","reassurance","none","digital_direct","mass","no","'Calculate coverage. Compare smarter.' Online calculator."),
    "DcTWqQPDa_e": ("brand_corporate","brand_corporate","other","pride","ratings_awards","none","sme_corporate","no","'Decade of Dedication', 10 Years of Excellence award."),
    "DcVULHvFxIK": ("brand_corporate","other","other","aspiration","none","none","youth","no","'Enableship 2026' - internship/programme teaser."),
    "DU-cDplgMZ1": ("individual_life","other","other","reassurance","none","none","women","no","Film still, woman in office. Message not readable at thumbnail size."),
    "DUvXTIGAams": ("individual_life","other","other","fear","none","none","mass","no","Film still, man in car at night. Message not readable at thumbnail size."),
    "DVN44h_ABrE": ("individual_life","other","other","aspiration","none","none","youth","no","Film still, three young men. Message not readable at thumbnail size."),

    # ---------------- IGI Life ----------------
    "Db-Jv0bDaPD": ("health","other","other","aspiration","none","none","affluent","no","Vitality: healthy meal, lifestyle imagery."),
    "Db2v7eQDbqi": ("health","price_affordability","convenience_digital","aspiration","none","digital_direct","youth","no","Vitality x foodpanda, PKR 500 weekly reward."),
    "Db7rzY9AVzh": ("individual_life","protection","other","reassurance","none","none","youth","yes","MYTH SERIES: 'I only need life insurance when I have a family.' Direct objection handling."),
    "DbvBkKNihue": ("individual_life","protection","other","reassurance","none","none","mass","yes","'LIFE INSURANCE MYTH' series title card."),
    "DcasP2TjOGX": ("individual_life","protection","other","reassurance","none","none","mass","no","'WHO ARE YOU PROTECTING TODAY? Because every stage of life brings something worth protecting.'"),
    "DcATDr3j3PY": ("brand_corporate","national_trust","other","pride","none","none","mass","no","14 August Independence Day."),
    "DcBMHQKo-v6": ("brand_corporate","national_trust","other","pride","none","none","mass","no","'Meri Pehchaan Pakistan' group photo."),
    "Dcdab3GFWHn": ("brand_corporate","brand_corporate","other","duty","none","none","mass","no","Hepatitis screening camp with Transparent Hands. CSR."),
    "DciNdoxjbDG": ("individual_life","other","other","fear","none","none","women","no","Woman stressed at desk. Part of the myth/objection series."),
    "DcLPa5CD6T0": ("child_education","education_marriage","other","reassurance","none","none","mass","yes","MYTH SERIES: 'My child's education is still years away. There's no need to plan for it yet.'"),
    "DcN-s7Moy5R": ("brand_corporate","brand_corporate","other","aspiration","none","none","youth","no","'Who's Most Likely To?' staff social content."),
    "DcN0QDWCNxB": ("health","other","other","reassurance","none","none","women","no","Vitality: screen time / wellbeing."),

    # ---------------- Askari Life ----------------
    "Dawwl8XlG62": ("child_education","education_marriage","protection","reassurance","none","call_centre","mass","no","'PLAN AHEAD WITH ASKARI LIFE EDUCATION PLAN'. #IraadaKaro. NB collides with State Life's 'Plan ahead' line."),
    "Db-rn1YNn8B": ("individual_life","protection","other","aspiration","none","call_centre","mass","no","'Ek chhotay se iraade se shuru hota hai'. Campaign platform."),
    "Db2YMNtk5rg": ("individual_life","protection","other","reassurance","none","call_centre","youth","yes","Chat UI: 'You added peace of mind / Uncertainty: Removed from chat'. Objection framing."),
    "DbH_qeKnBPu": ("child_education","education_marriage","other","duty","none","call_centre","mass","no","'SCHOOL FEES TODAY. University dreams tomorrow.'"),
    "Dbk1r-mnNNk": ("savings_investment","savings_return","other","aspiration","none","digital_direct","mass","no","'BONUS ADDED - insurance notification you don't expect'. Golden Bonus Rewards."),
    "DbYlN_GnLR2": ("individual_life","other","other","reassurance","none","call_centre","youth","no","'3 things the rain teaches us about life'."),
    "DcdED9xHHID": ("savings_investment","savings_return","other","aspiration","none","digital_direct","mass","no","'What if your plan rewarded you from year 1?' Golden Path."),
    "DcIqpNInO6e": ("individual_life","other","other","aspiration","none","call_centre","mass","no","'Aapki window mein kaunsa khwab hai?' - apartment windows, aspirations in Urdu."),
    "DcS3gNbErb4": ("individual_life","protection","savings_return","aspiration","none","call_centre","mass","no","'YOUR FUTURE DESERVES MORE THAN JUST A PLAN'."),
    "DKymg91NZ_o": ("individual_life","other","other","aspiration","none","call_centre","mass","no","Hand/phone creative, message not readable at thumbnail size."),
    "DKymitONopz": ("individual_life","other","other","aspiration","none","call_centre","youth","no","'YOUR FUTURE is calling' - phone call UI."),
    "DKymkdrNBgN": ("brand_corporate","brand_corporate","other","aspiration","none","call_centre","mass","no","'IRAADA KARO - Jeeto ka, har ghari'. Platform end-frame."),

    # ---------------- TPL Life ----------------
    "DauMnaKgFfh": ("health","protection","other","aspiration","none","digital_direct","affluent","no","Globewell: 'Global adventures deserve global protection', 185+ countries."),
    "Db2NdLgCO4g": ("health","protection","convenience_digital","reassurance","none","digital_direct","affluent","no","'Your health, insured worldwide', 24/7 medical assistance."),
    "Db7pvpIFvNs": ("brand_corporate","faith_compliance","other","pride","none","none","sme_corporate","no","MobiLink/Bank partnership signing, 'Partnering for greater protection'."),
    "DbAO2HPgPVi": ("health","protection","other","aspiration","none","digital_direct","affluent","no","'Vacationing abroad? Globewell travels with you', $500,000."),
    "DbpYV5UDu4L": ("brand_corporate","brand_corporate","other","pride","none","none","sme_corporate","no","'Celebrating a successful H1 2026' with DIB."),
    "DbSMK7XCNHi": ("health","protection","other","reassurance","none","digital_direct","affluent","no","'World-class hospitals. One international plan.' $500,000."),
    "DcAifPGiqQj": ("brand_corporate","national_trust","other","pride","none","none","mass","no","79 Years Independence stamp illustration."),
    "DcK1y_ngGU4": ("health","protection","other","aspiration","none","digital_direct","affluent","no","'From Dubai to London, Globewell has got your back'."),
    "DY04PKAiNdu": ("brand_corporate","faith_compliance","other","reassurance","none","none","mass","no","Eid-ul-Adha greeting."),
    "DYWPP3_ACKa": ("health","protection","other","reassurance","none","digital_direct","affluent","no","'Global care. Zero compromise.' $500,000."),
    "DZHJYUpgGo4": ("health","price_affordability","protection","reassurance","none","digital_direct","mass","no","Sehat Zindagi: 'Quality healthcare doesn't have to be expensive', PKR 1,250,000."),
    "DZUBKHvCK9b": ("health","price_affordability","other","reassurance","none","digital_direct","rural","no","'Affordable plan. Nationwide coverage.' Top-tier hospitals."),

    # ---------------- Pak-Qatar Family Takaful ----------------
    "Db-9UNvjBoR": ("brand_corporate","brand_corporate","other","pride","expert_authority","none","sme_corporate","no","CEO panellist, ICMA Future of Finance 2026. AA/AM2+ badge."),
    "Db5B8ZcDDwK": ("retirement","national_trust","faith_compliance","pride","sovereign_guarantee","none","sme_corporate","no","Appointed by Federal Government as Authorized Pension Fund Manager. Strongest state-linked claim by a private player."),
    "DbAYGZ1DF0h": ("savings_investment","savings_return","faith_compliance","aspiration","ratings_awards","digital_direct","mass","no","'THE SECOND PAYCHEQUE' - Mahana Bachat & Takaful Flexi Plan."),
    "DbDTAzQDIBc": ("brand_corporate","brand_corporate","other","pride","ratings_awards","none","sme_corporate","no","Best Bancatakaful Award, IAP-PSOA conference."),
    "DbKhUFcPC8Z": ("individual_life","other","other","reassurance","ratings_awards","none","mass","no","Man in interior, message not readable at thumbnail size."),
    "DbUy7M5DMow": ("health","other","other","duty","expert_authority","none","mass","no","World Hepatitis Day - 'the silent disease needs a loud voice'. CSR."),
    "DbvX7bhPajD": ("brand_corporate","other","other","reassurance","ratings_awards","none","mass","no","Near-blank creative; not readable."),
    "DcaX5R-CNB3": ("retirement","retirement","faith_compliance","reassurance","ratings_awards","agent","mass","no","'RETIREMENT DESERVES CERTAINTY - LIFETIME KAFALAT PLAN. A guaranteed halal pension for life.' Closest competitor claim to a guarantee."),
    "DcdG-TJiCOi": ("brand_corporate","brand_corporate","other","pride","none","none","mass","no","Bahawalpur branch inauguration."),
    "DciELVJiOxw": ("group_corporate","protection","other","duty","ratings_awards","agent","sme_corporate","no","'In uncertain moments, support matters most' - Group Health Takaful."),
    "DcOOz4CjJ_4": ("retirement","retirement","savings_return","aspiration","ratings_awards","agent","sme_corporate","no","'BUILD TOMORROW FROM TODAY' - Voluntary Pension Scheme, 20% tax credit."),
    "DcQs_BUjKCa": ("savings_investment","savings_return","faith_compliance","aspiration","ratings_awards","agent","mass","no","'SMALL STEPS. BIG FUTURE.' Mahana Bachat, monthly halal profit."),

    # ---------------- Dawood Family Takaful ----------------
    "DJ0_ugxt7-x": ("brand_corporate","convenience_digital","other","reassurance","none","digital_direct","mass","no","'PAY ONLINE. STAY COVERED.'"),
    "DJ_kjMmtmkh": ("takaful","faith_compliance","other","duty","none","none","mass","yes","'Takaful or Insurance? CHOOSE THE HALAL WAY!' Direct faith objection handling."),
    "DJbGZ9lodTR": ("brand_corporate","faith_compliance","other","reassurance","none","none","mass","no","Quranic verse, Muharram."),
    "DJIsQMStTvl": ("brand_corporate","national_trust","other","pride","none","none","mass","no","'PAKISTAN'S PRIDE' - armed forces salute."),
    "DJqqGDsNdIO": ("savings_investment","savings_return","faith_compliance","aspiration","none","none","mass","no","'SECURE TOMORROW, TODAY WITH SAHULAT PLAN'."),
    "DJWnn_7oIze": ("brand_corporate","national_trust","other","pride","none","none","mass","no","Air force / Urdu verse patriotic post."),
    "DKcJknbtstL": ("takaful","protection","faith_compliance","reassurance","none","none","mass","yes","'TAKAFUL COVERS MORE THAN YOU THINK - it's a full-circle protection'. Category education."),
    "DKg6sSpNq_k": ("takaful","protection","faith_compliance","reassurance","none","digital_direct","mass","no","'SARMAYA TAKAFUL PLAN - hassle-free withdrawal, 100% Shariah-compliant'."),
    "DKHlC7_tjkx": ("child_education","education_marriage","faith_compliance","aspiration","none","none","mass","no","'SAMAR TAKAFUL - secure a bright future for education & marriage'."),
    "DKLt2M6v0Dc": ("brand_corporate","national_trust","other","pride","none","none","mass","no","Youm-e-Takbeer, 28 May 1998."),
    "DKRl67-tiN4": ("takaful","faith_compliance","protection","duty","expert_authority","none","mass","yes","'TIE YOUR CAMEL AND TRUST IN ALLAH' - Prophet Muhammad (PBUH), Jami at Tirmidhi 2517. Answers the fatalism objection with scripture. Sharpest objection-handling ad in the whole sample."),
    "DKUmrwSOwlY": ("brand_corporate","national_trust","other","pride","celebrity","none","mass","no","Congratulating Arshad Nadeem, Asian Athletics gold."),

    # ---------------- EFU Life Window Takaful / secondary EFU ----------------
    "Da5c7P7s9ny": ("brand_corporate","convenience_digital","other","reassurance","none","digital_direct","mass","no","'FAST. EASY. DIGITAL.' Premium payments on WhatsApp."),
    "DaAbyhfsHKW": ("unclear","other","other","reassurance","none","none","mass","no","Video post, poster frame did not load."),
    "DaAcbOdDHuo": ("takaful","savings_return","faith_compliance","aspiration","none","digital_direct","mass","no","Hemayah: 'Your lifestyle is the result of your hard work. Want to make it last?'"),
    "DaYLFMM81j": ("unclear","other","other","reassurance","none","none","mass","no","Video post, poster frame did not load."),
    "DaZZPsMq9M": ("unclear","other","other","reassurance","none","none","mass","no","Video post, poster frame did not load."),
    "DaK2RcqDMW6": ("brand_corporate","brand_corporate","convenience_digital","pride","expert_authority","none","sme_corporate","no","CEO: 'REDEFINING LEADERSHIP IN INSURANCE... securing the future of millions of Pakistanis'."),
    "Damvw69uB1A": ("takaful","faith_compliance","convenience_digital","reassurance","none","none","rural","no","Hemayah x EFU General x BankIslami - micro & nano takaful."),
    "DbM8kFHMcF6": ("brand_corporate","other","other","duty","none","call_centre","mass","no","'HEAVY RAIN ADVISORY - stay safe, stay prepared'. Public-service post."),
    "DbuoyOFsZIx": ("retirement","national_trust","retirement","reassurance","sovereign_guarantee","none","mass","no","'STRENGTHENING PAKISTAN'S RETIREMENT FUTURE WITH GOVERNMENT INITIATIVES' - Federal Hemayah Pension Fund. A private insurer borrowing state association."),
    "DbupP4oM3Md": ("health","convenience_digital","other","reassurance","none","digital_direct","mass","no","'A complete healthcare and wellbeing experience' via EFU Life mHealth."),
    "DcA0mhrONio": ("brand_corporate","brand_corporate","other","pride","heritage_scale","none","mass","no","'Born from a DREAM' - heritage/founding story."),
    "DcDdjWLjnQY": ("health","convenience_digital","other","reassurance","none","digital_direct","sme_corporate","no","'Join EFU Life mHealth - a digital healthcare platform for modern medical consultants'."),

    # ---------------- UBL Fund Managers ----------------
    "C0MOKhSocyS": ("savings_investment","convenience_digital","other","reassurance","none","digital_direct","mass","no","UBL Funds digital mobile app access."),
    "Dca-Uw3grJ2": ("savings_investment","convenience_digital","other","aspiration","none","digital_direct","mass","no","'Faisla yaqeen ka' campaign line, event booth."),
    "DcaoSS1s1DR": ("brand_corporate","national_trust","other","pride","none","none","mass","no","'Spirit of Pakistan' 14 Aug cake."),
    "Dcazxa3jPIk": ("savings_investment","convenience_digital","savings_return","reassurance","ratings_awards","digital_direct","affluent","no","'Fauri Paisa - your money, when you need it. Instant redemption up to PKR 1,000,000.' Liquidity as the claim."),
    "DcdL47_jIrP": ("brand_corporate","brand_corporate","other","pride","none","none","sme_corporate","no","MOU signing ceremony."),
    "DceOx12CGcI": ("brand_corporate","faith_compliance","other","reassurance","none","none","mass","no","Rabi ul Awal / Eid Milad-un-Nabi greeting."),
    "DcLgeH8ILnr": ("brand_corporate","brand_corporate","other","aspiration","none","none","sme_corporate","no","'We're Hiring' - Relationship Manager. Recruitment."),
    "DcNc9TVjH77": ("brand_corporate","brand_corporate","other","aspiration","none","none","sme_corporate","no","'We're Hiring' - Portfolio Manager. Recruitment."),
    "DcNdT6qjOCZ": ("brand_corporate","brand_corporate","other","aspiration","none","none","sme_corporate","no","'We're Hiring' - Assistant Portfolio Manager. Recruitment."),
    "DcNmmoKjK0L": ("savings_investment","savings_return","other","aspiration","none","none","mass","no","'Money Matters' wealth expo, Gilgit, with Al-Ameen Funds."),
    "DLj9FzXl_3H": ("savings_investment","convenience_digital","other","reassurance","none","digital_direct","mass","no","'Invest via RAAST - tap, transfer, grow'."),
    "DNkisDyMwXu": ("savings_investment","convenience_digital","other","aspiration","none","digital_direct","mass","no","Al-Ameen rebrand: 'smarter, better... the new us'."),

    # ---------------- Al Meezan Investments ----------------
    "Da6_VMcjBTw": ("savings_investment","savings_return","other","reassurance","expert_authority","none","affluent","no","'Investors curious on P/E re-rating' - KSE100 analysis chart. Genuinely analytical content."),
    "Dbd2rDrEYya": ("savings_investment","savings_return","other","reassurance","expert_authority","none","affluent","no","'Market Outlook 2026' with CEO Imtiaz."),
    "DcaT_zvjFTc": ("savings_investment","savings_return","faith_compliance","aspiration","none","digital_direct","mass","no","'HOME BUILDER PLAN - some doors open long before you reach them'."),
    "Dcdq4IbNeuI": ("savings_investment","faith_compliance","convenience_digital","reassurance","ratings_awards","digital_direct","mass","no","Urdu: easy and fully Shariah-compliant investing. AM1 badge."),
    "Dcf7OwnjP6V": ("brand_corporate","brand_corporate","other","aspiration","none","none","youth","no","Iqra University Career Fair. Recruitment."),
    "DcFqcCxEgUF": ("brand_corporate","brand_corporate","other","duty","none","none","mass","no","'Al Meezan Cares' public service message."),
    "DcfTi9zijsM": ("brand_corporate","faith_compliance","other","reassurance","none","none","mass","no","Eid Milad-un-Nabi, Green Dome."),
    "Dcik9WYjupP": ("brand_corporate","national_trust","other","pride","none","none","mass","no","Landscape / patriotic imagery."),
    "DcIqvkLFGD5": ("savings_investment","convenience_digital","other","reassurance","none","digital_direct","mass","no","'How to generate RAAST Investment ID' on desktop."),
    "DcLPygWAXyz": ("savings_investment","convenience_digital","other","reassurance","none","digital_direct","mass","no","'How to generate RAAST Investment ID' on mobile."),
    "DcNvAhKDByh": ("savings_investment","savings_return","other","aspiration","none","none","mass","no","'WHERE MONEY MEETS MEANING' - Money Matters expo silver partner."),
    "DcTGy3DFNJi": ("savings_investment","faith_compliance","other","humour","none","digital_direct","youth","yes","'MYTH & FACT' with 'Al Meezan Buddy' robot mascot. Objection handling."),
}


def main() -> None:
    wb = openpyxl.load_workbook(AUDIT / "capture.xlsx")
    ws = wb["Ads"]

    # clear the provisional rows written before the creative was in hand
    for r in range(2, ws.max_row + 1):
        for col in "BDEFGHIJKLMNOPQRSTUV":
            ws[f"{col}{r}"] = None

    with (AUDIT / "posts.csv").open(encoding="utf-8", newline="") as fh:
        posts = list(csv.DictReader(fh))

    cols = "BDEFGHIJKLMNOPQRSTUV"
    row = 2
    coded = unclear = 0
    for p in posts:
        sc = p["shortcode"]
        c = C.get(sc)
        if c:
            product, cl1, cl2, reg, proof, cta, aud, obj, note = c
            coded += 1
        else:
            product, cl1, cl2, reg, proof, cta, aud, obj = (
                "unclear", "other", "other", "reassurance", "none", "none", "mass", "no")
            note = "NOT YET CODED - open the full-size file in media/ and code it."
            unclear += 1
        if p.get("reposted_from"):
            note = f"REPOST from @{p['reposted_from']}. " + note
        vals = [
            p["brand"], "meta", "", "2026-08-27",
            "video" if p["kind"] == "reel" else "static", None, "mixed",
            product, cl1, cl2, reg, proof, cta, aud, obj, 1, None,
            p["source_url"], "claude", (note + f" [img: {p['image_file']}]")[:500],
        ]
        for col, v in zip(cols, vals):
            ws[f"{col}{row}"] = v
        row += 1

    wb.save(AUDIT / "capture.xlsx")
    print(f"{row-2} rows written: {coded} coded from the creative, {unclear} left unclear.")


if __name__ == "__main__":
    main()
