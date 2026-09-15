# -*- coding: utf-8 -*-
"""Fill capture.xlsx Ads sheet from posts.csv + hand-coding read off the contact sheets."""
import csv, re, openpyxl
from datetime import datetime as _dt
from pathlib import Path

SCAN = Path("scan")

# post_id -> (product_focus, gen, content_type, claim1, claim2, register, proof, frame,
#             production, eco, platform, objection, lang, notes)
C = {
# ---------------- HBL ----------------
"DcaWojblXMd":("HBL Credit Card x Mercantile","geny","tactical_promo","lifestyle_reward","price_affordability","aspiration","none","none_people","template_graphic","no","","no","english_first","Collab post w/ Mercantile Pakistan. 'what's next will be more within reach' - iPhone launch teaser. Bank is the payment rail, not the story."),
"DcbDuYkJ_f_":("HBL CarLoan","geny","product_feature","credit_access","price_affordability","aspiration","claim_statistics","none_people","template_graphic","no","","no","english_first","'Mark-up rate 11.99% fixed for Tenure / Insurance 1.4% / Free Tracker'. Rate printed on the artwork - rare in this set."),
"DcdmVI7Pe3o":("HBL Shop Now Pay Later","geny","tactical_promo","price_affordability","","aspiration","none","none_people","template_graphic","no","Shop Now, Pay Later","no","english_first","'The coolest upgrade you'll make this summer' - AC instalments at 0%."),
"Dcdzgr2PcdY":("HBL Business Debit Card","geny","product_feature","convenience_digital","","aspiration","none","customer","stock","no","","no","english_first","'Make Every Move Count' - woman at laptop, Notion/Trello/Monday.com logos. SME-freelancer framing."),
"DciKpAvjoPq":("She's Next (Visa x HBL)","geny","brand_building","financial_inclusion","brand_corporate","pride","none","customer","template_graphic","no","She's Next","no","english_first","'Where women connect, ideas grow - She's Next Club is back'. Only women-in-business platform in the set."),
"DcirADLDiWd":("She's Next (Visa x HBL)","unclear","brand_building","financial_inclusion","","aspiration","none","none_people","template_graphic","no","She's Next","no","english_first","'She's Next Club is coming online' - handwritten-note device."),
"DcLmhpYCFqb":("HBL Islamic / Hajj","mixed","brand_building","faith_compliance","national_trust","duty","expert_authority","customer","studio","no","Safar-e-Hajj","no","urdu_first","Urdu lead: 'safar-e-Hajj mein aap ka khidmatguzar'. Badge: EXCLUSIVE HAJJ BANKING PARTNER, Ministry of Religious Affairs. Only faith territory a conventional bank occupies here."),
"DcYO6JiIX2b":("HBL Islamic / Hajj","mixed","brand_building","faith_compliance","national_trust","duty","expert_authority","none_people","studio","no","Safar-e-Hajj","no","urdu_first","Reel cut of the Hajj key visual."),
"DcYpwmbBr2q":("HBL Islamic / Hajj","mixed","brand_building","faith_compliance","national_trust","duty","expert_authority","none_people","studio","no","Safar-e-Hajj","no","urdu_first","Second reel cut of the same Hajj key visual - same asset run twice within a day."),
"DcV3W87igQ2":("HBL careers (The League)","genz","recruitment","brand_corporate","","pride","none","staff","studio","no","The League / #LeadTheNew","no","english_first","'THE LEAGUE 2026 HAS ARRIVED - Meet the MTs ready to #LeadTheNew at HBL'."),
"DWytximGBUX":("brand","unclear","brand_building","brand_corporate","","reassurance","none","none_people","studio","no","","no","english_first","Macro shot of the HBL logo on teal fabric - brand film, no proposition legible."),
"DZKh7n-ucE7":("brand","unclear","corporate_pr","brand_corporate","","pride","expert_authority","executive","studio","no","","no","english_first","'Successful Brands that Create Impact' - talking-head executive interview."),
# ---------------- UBL ----------------
"DbLRyZkTA7c":("UBL corporate / nation-building","unclear","corporate_pr","national_trust","","pride","claim_statistics","none_people","template_graphic","no","Investing in Our Land","no","english_first","'Commitment of 40 billion Rupees': PKR 10bn New University, 22bn Khushhali Bank equity, 8bn agriculture subsidy. Largest hard number anywhere in the scan."),
"DcibBB3ikhY":("Raast","geny","product_feature","convenience_digital","","reassurance","none","customer","stock","yes","Raast","no","english_first","'FAST, SECURE, AND HASSLE-FREE!' Raast lockup carried larger than any UBL product mark."),
"DcieNDuCrjK":("Raast","geny","product_feature","convenience_digital","","reassurance","none","customer","stock","yes","Raast","no","english_first","'SWIFT AND SIMPLE PAYMENTS!' - same layout, different stock model."),
"DciGQjnCD5L":("UBL Digital App / Raast RTP","geny","product_feature","convenience_digital","","reassurance","none","none_people","template_graphic","yes","Raast","no","english_first","'Raast is now the new way to shop' - Request to Pay; app version 3.41.2 printed on the artwork."),
"Dcihy4_FQ6N":("Raast OTC","mixed","product_feature","financial_inclusion","convenience_digital","reassurance","none","customer","stock","no","Raast","yes","english_first","'Raast over the counter (OTC) available at all branches' - 'No charges at all'. The one post addressing a non-app customer."),
"DciND4fFWB7":("Raast","geny","product_feature","convenience_digital","","reassurance","none","none_people","stock","yes","Raast","no","english_first","'YOUR MOBILE NUMBER IS NOW YOUR ACCOUNT NUMBER WITH RAAST!'"),
"DcirlvLCiGJ":("UBL Cards discounts","geny","tactical_promo","lifestyle_reward","price_affordability","aspiration","none","none_people","stock","no","","no","english_first","'Gather Around Great Food - 15% off at LA Serre Restaurant' - single outlet, Bahawalpur only."),
"DciUQCHFVpW":("Raast","geny","product_feature","convenience_digital","","reassurance","none","customer","stock","yes","Raast","no","english_first","'SEAMLESS TRANSFER, ZERO HASSLE!'"),
"DcivammDRkc":("UBL Debit Card / Google Wallet","geny","tactical_promo","convenience_digital","lifestyle_reward","aspiration","none","none_people","template_graphic","yes","Tap. Pay. Win","no","english_first","'Every PKR 3,000 spent in-store or online earns one lucky draw entry' - monthly smartwatch."),
"DciVbCMlXz4":("UBL Digital App","geny","product_feature","security_fraud","convenience_digital","reassurance","none","none_people","template_graphic","yes","","yes","urdu_first","Urdu lead: 'apna account fori freeze karein' - self-service block/unblock. Only proactive fraud-control feature in 84 posts."),
"DciWGN3iOfJ":("UBL Digital Buddy","geny","tactical_promo","convenience_digital","price_affordability","aspiration","none","customer","stock","yes","Digital Buddy","no","english_first","'Get A Chance to Win Up to PKR 500 Cashback' - WhatsApp registration, offer valid to 30 August."),
"DciXQg5leXy":("Raast","geny","product_feature","convenience_digital","","reassurance","none","customer","stock","yes","Raast","no","english_first","'QUICK TRANSFERS, INSTANT PAYMENTS - Only with Raast, the Best Payment Network for Digital Transactions'. UBL advertising the state rail, not itself."),
# ---------------- Meezan ----------------
"DcgK8ZiFIpZ":("Meezan Mobile App / Payoneer","geny","product_feature","convenience_digital","","aspiration","none","none_people","template_graphic","yes","","no","english_first","'WORK GLOBAL, WITHDRAW LOCAL' - Payoneer credited into the Meezan app. Freelancer segment."),
"DcgpSRwgYZM":("Meezan customer protection","mixed","csr","security_fraud","","fear","none","none_people","stock","no","Scam Alert","yes","english_first","'FAKE PARCEL SCAM ALERT' - one of two scam advisories in twelve posts."),
"Dch_IiNDRF1":("Meezan Student Debit Card","genz","tactical_promo","lifestyle_reward","","aspiration","none","none_people","template_graphic","no","","no","english_first","'Study Smart. Play Hard. Save More.' - up to 40% off at twelve padel courts."),
"DciEKVPHHf7":("Meezan Visa Cards","geny","tactical_promo","lifestyle_reward","","aspiration","none","none_people","template_graphic","no","Spend & Win","no","english_first","'SPEND & WIN - use your card internationally and win big': Honda CD70 / iPhone 17 by spend tier."),
"DciewUhAeS6":("brand","unclear","corporate_pr","brand_corporate","","pride","ratings_awards","none_people","template_graphic","no","Six-Time Best Bank","no","english_first","'SIX-TIME BEST BANK 2018, 2020, 2023, 2024, 2025 & 2026 - Two wins. One incredible achievement.'"),
"Dcilb5oD_D0":("Meezan COII","genx","product_feature","savings_return","faith_compliance","aspiration","claim_statistics","none_people","template_graphic","no","","no","english_first","'MAKE TODAY'S SAVINGS COUNT FOR TOMORROW - expected return of 10% per annum', COII 1.5 years, Shariah-compliant. Only headline return rate in the set."),
"DcinYAUFkQE":("Meezan customer protection","mixed","csr","security_fraud","","fear","none","none_people","template_graphic","no","","yes","english_first","'STAY VIGILANT AND PROTECT YOURSELF AGAINST IMPERSONATION!'"),
"DcivvLEAFza":("Meezan Premium Banking / RDA","genx","product_feature","premium_status","savings_return","aspiration","claim_statistics","none_people","template_graphic","no","Meezan Premium","no","english_first","Qualification thresholds printed on the artwork: avg PKR 3m accounts / 5m term deposits / 10m combined / 8m Roshan Apna Ghar."),
"DciXIpBjI5W":("Meezan Senior Citizen Account","boomer_plus","product_feature","life_stage_care","protection","reassurance","none","customer","stock","no","","no","english_first","'A LIFETIME OF EXPERIENCE DESERVES CARE' - elderly hands on a cane. The only post in 84 aimed at over-60s."),
"DcQXtSJnO_E":("brand","unclear","corporate_pr","brand_corporate","","pride","none","executive","event_photo","no","","no","english_first","Ribbon-cutting: 'Inauguration Ceremony of Meezan Bank's First Round-the-Clock Meezan Service Center at Clifton Bridge, Karachi'."),
"DcT0YMtnM3v":("brand","unclear","corporate_pr","brand_corporate","","pride","ratings_awards","executive","event_photo","no","Six-Time Best Bank","no","english_first","Executives receiving the Pakistan Banking Award trophy."),
"DcTxpZsjBgV":("brand","unclear","corporate_pr","brand_corporate","","pride","ratings_awards","none_people","template_graphic","no","Six-Time Best Bank","no","english_first","'Best Bank of Pakistan FOR THE 6TH TIME' - trophy render. Third award post inside the same window."),
# ---------------- Bank Alfalah ----------------
"Dah4IWWIuw9":("Alfalah Visa Credit Card","geny","tactical_promo","price_affordability","","aspiration","none","none_people","template_graphic","no","","no","english_first","'Spend Today, Pay Over Time, Travel Without Compromise - 0% Markup, 3 Easy Instalments with Low Processing Fee'."),
"Dbp8GYJswi1":("Alfalah Cards (POS & e-commerce)","geny","tactical_promo","lifestyle_reward","","aspiration","none","customer","stock","no","Travel More, Spend Smarter","no","english_first","Woman with luggage against Big Ben - 'Spend Smarter Internationally, Win Big'."),
"DbyDbWrIjzK":("brand","unclear","brand_building","brand_corporate","","reassurance","none","none_people","template_graphic","no","The Way Forward","no","english_first","A video whose entire frame is the logo lockup on white. No proposition, no product, no person."),
"DcdaAJboi6B":("Bank Alfalah Partner (WhatsApp)","geny","product_feature","convenience_digital","","reassurance","none","customer","stock","yes","Partner is always on","no","english_first","'WhatsApp Hi to 021 111 225 111 to access Bank Alfalah Partner'."),
"DcdLj-QoI1O":("Alfalah SBS Instalments (TCL)","geny","tactical_promo","price_affordability","","aspiration","claim_statistics","none_people","template_graphic","no","SBS Instalments","no","english_first","A full eight-column instalment price table for TCL LEDs, 3-24 months, on a social tile. Unreadable at feed size."),
"DcijhtSIpwC":("Alfalah Credit Card e-statements","geny","product_feature","convenience_digital","","reassurance","none","none_people","stock","no","","yes","english_first","'CREDIT CARD E-STATEMENTS, AT YOUR FINGERTIPS - SMS Green to 8287'."),
"DciLij6Ii25":("Alfalah auto-debit","geny","product_feature","convenience_digital","","reassurance","none","none_people","template_graphic","no","","yes","english_first","'Forgot to pay your Bill AGAIN?' - enroll your credit card for auto debit. Names the actual customer failure."),
"DciofU3ISry":("Alfalah Visa Platinum","genx","product_feature","premium_status","lifestyle_reward","aspiration","none","none_people","template_graphic","no","","no","english_first","'Where Luxury Travels With You' - lounge access, travel rewards, purchase protection, 200+ merchants."),
"Dcir-kRIKru":("Alfalah customer protection","mixed","csr","security_fraud","","fear","none","none_people","template_graphic","no","Think Before You Pay","yes","english_first","'Fraud Alert: Fake Driving License Offers Without Test' - two columns of body copy, How You Get Trapped / How To Protect Yourself."),
"DciTKqsoSE4":("Alfalah Supplementary Credit Card","geny","product_feature","credit_access","","reassurance","none","customer","stock","no","","no","english_first","'Empower your loved ones with Bank Alfalah Supplementary Credit Card - Share your credit limit'."),
"DcixA2yKYkW":("Alfalah SBS Instalments (American General)","geny","tactical_promo","price_affordability","","aspiration","claim_statistics","none_people","template_graphic","no","SBS Instalments","no","english_first","Second appliance price-table tile inside three days."),
"DciXctCiOjH":("Alfalah owned channels","geny","corporate_pr","brand_corporate","","reassurance","none","none_people","template_graphic","yes","","no","english_first","'NEVER MISS AN UPDATE THAT MATTERS - follow us on WhatsApp, Facebook, Instagram'. A house ad for the feed itself."),
# ---------------- JS Bank ----------------
"DcgAg0Kjaf_":("JS Solar Finance","genx","product_feature","credit_access","sustainability","reassurance","none","none_people","template_graphic","no","Zindagi Karo Roshan","yes","other","Roman Urdu eligibility checklist: 'Kya aap active taxpayer hain? / credit history tasalli-bakhsh hai? / umar 65 saal sey kam hai?' Green Climate Fund mark."),
"DcgAk8ZjTCL":("JS Solar Finance","mixed","brand_building","sustainability","credit_access","duty","none","customer","studio","no","Zindagi Karo Roshan","no","other","Night-lit clinic interior, commissioned photography. The only genuinely art-directed documentary imagery in the scan."),
"DcgAol-DRjU":("JS Solar Finance","genx","brand_building","credit_access","sustainability","duty","none","none_people","template_graphic","yes","Zindagi Karo Roshan","no","other","'PATIENTS KA BHAROSA BARQARAR - apnay clinic ko rakhein powered bina interruption ka JS Solar Finance kay saath. Zindagi Karo Roshan!'"),
"DZDIkz3jVBP":("JS Solar Finance","genx","brand_building","sustainability","credit_access","duty","none","none_people","template_graphic","no","Zindagi Karo Roshan","no","english_first","'ROSHAN SCHOOL ROSHAN KAL with JS Solar Finance'."),
"DZDIFC9PYAu":("JS Solar Finance","mixed","brand_building","sustainability","","duty","none","customer","studio","no","Zindagi Karo Roshan","no","other","Family in a lit room - same documentary treatment as the clinic."),
"DZDH-XnDX-C":("JS Solar Finance","genx","product_feature","sustainability","credit_access","duty","none","none_people","template_graphic","no","Zindagi Karo Roshan","no","other","Roman Urdu: 'JS Solar Finance sey school facilities ko behtar banayein aur students kay liye roshan mustaqbil ki bunyaad rakhein.'"),
"DcgZtcFDc5H":("JS Credit Card","geny","tactical_promo","price_affordability","","aspiration","none","none_people","template_graphic","no","Power Up with Cashback","no","english_first","'Pay your electricity bill with your JS Credit Card and enjoy 5% Cashback'."),
"DcgZwWGDcr8":("JS Credit Card","geny","tactical_promo","price_affordability","","humour","none","customer","studio","no","Power Up with Cashback","no","english_first","Art-directed orange set, man reacting to a 5% cashback bill. A commissioned shoot for a tactical offer."),
"DcgZzIGDTpA":("JS Credit Card","geny","tactical_promo","price_affordability","","aspiration","none","none_people","template_graphic","no","Power Up with Cashback","no","english_first","'POWER UP WITH 5% CASHBACK' type-led tile."),
"DciCE8ujfu_":("JS Her Debit Card","geny","tactical_promo","lifestyle_reward","","aspiration","none","customer","stock","no","JS Her","no","english_first","'STAY FIT, SAVE MORE. WITH JS HER DEBIT CARD, ENJOY 20% OFF AT CLIFTON COURTYARD' - the athlete cast on a women's product is male."),
"DciCJX_Da79":("JS Her Debit Card","geny","tactical_promo","lifestyle_reward","","aspiration","none","none_people","stock","no","JS Her","no","english_first","'STYLE, SHOP & SAVE ... ENJOY 40% OFF AT PENG' salon and spa."),
"DciCMSJDRvH":("JS Her Debit Card","geny","tactical_promo","lifestyle_reward","","aspiration","none","none_people","template_graphic","no","JS Her","no","english_first","'A DAY OF FUN FOR KIDS ... 20% OFF AT GIGGLE TOWN'. Illustrated children, no adult woman in frame on a women's product."),
# ---------------- Easypaisa ----------------
"Da8MI_ZKrjK":("easypaisa debit card","genz","brand_building","lifestyle_reward","","humour","none","customer","ugc","no","","no","english_first","Creator meme: 'when someone is yelling stop spending so much on coffee but my name is easypaisa debit card and now I can't stop'."),
"Db0u7zKHWvW":("easypaisa app","genz","product_feature","security_fraud","convenience_digital","humour","none","customer","ugc","yes","","yes","english_first","'guess which sound is legit' - the payment-confirmation sound as a fraud test. Phone-shot, vertical."),
"Db2z0YalTVc":("easypaisa app","geny","product_feature","price_affordability","","reassurance","none","none_people","template_graphic","yes","Did you know?","yes","english_first","'All funds received in easypaisa are free-of-charge and will remain free-of-charge.' A pricing promise stated flatly."),
"Dba_RSLEjx3":("easypaisa QR","genz","brand_building","convenience_digital","","humour","none","customer","ugc","yes","","no","english_first","'I think you're pretty ... obsessed with QR'. Creator-shot, phone-native."),
"DbIxiQzqN4Q":("none","genz","brand_building","other","","humour","none","celebrity","ugc","no","","no","english_first","Borrowed Haaland meme, 'us waiting to see who comments nice first'. No product, no proposition - engagement bait."),
"Dblf4Iymt4U":("easypaisa customer protection","genz","csr","security_fraud","","humour","none","none_people","ugc","no","","yes","english_first","'sitting down for the documentary because I wasn't today's scammer success story' - a fraud warning delivered as a meme."),
"DcAgxFfF70X":("brand","mixed","calendar_topical","national_trust","","pride","none","none_people","template_graphic","no","","no","english_first","'Pakistan, with love. From the dreams of 1947 to the hopes of tomorrow. Happy 79th Independence Day.'"),
"DcgkW-jKW7l":("easypaisa term deposit / savings","genz","product_feature","savings_return","","humour","none","none_people","template_graphic","yes","gold star behaviour","no","english_first","'gold star behaviour: auto-investing with term deposit instead of impulse-buying'. Savings sold in meme language - unique in the set."),
"DcJc2X7iq04":("easypaisa debit card","genz","brand_building","lifestyle_reward","","humour","none","customer","ugc","no","","no","english_first","'kinda cool that one card gets you discounts at all the places you love'."),
"DcWNoKSKB_W":("easypaisa debit card","genz","brand_building","other","","humour","none","staff","ugc","no","","no","english_first","'I am tired of making content. Can you all just order easypaisa debit cards?' - the social team on camera as the brand."),
"DTkvogbCjZy":("easypaisa debit card","geny","product_feature","price_affordability","convenience_digital","aspiration","none","customer","studio","no","har ATM apna","yes","urdu_first","'har ATM apna - zero ATM fee with easypaisa debit cards'. Removes a named cost objection."),
"DYSMJN6sl5F":("easypaisa customer protection","genz","csr","security_fraud","","humour","none","customer","ugc","no","","yes","other","Roman Urdu scam bubbles: 'trading / committee / nai scheme / plaat barai farokht'."),
# ---------------- SadaPay ----------------
"Da5zUWAxGWY":("SadaPay debit card","geny","brand_building","convenience_digital","lifestyle_reward","aspiration","celebrity","none_people","studio","no","","no","english_first","'Your perfect travel companion' - card held against a beach town; Shoaib Akhtar Founder's Club framing in the reel."),
"Da71cGwsPcx":("SadaPay app (hotels)","geny","product_feature","convenience_digital","","reassurance","none","customer","studio","yes","","no","english_first","'Travelling? Secure your hotel room in seconds.' Commissioned lifestyle shoot."),
"Da721Ohs-2J":("SadaPay Gamezone","genz","product_feature","lifestyle_reward","","humour","none","none_people","template_graphic","yes","Gamezone","no","english_first","'INTRODUCING SADAPAY GAMEZONE' in pixel-art type."),
"DanegQisPn0":("SadaPay app (flights)","geny","product_feature","convenience_digital","","aspiration","none","customer","studio","yes","","no","english_first","'Book international flights instantly. No hassle.'"),
"Dax5v-Bsvog":("SadaPay app (experiences)","geny","product_feature","convenience_digital","lifestyle_reward","aspiration","none","customer","studio","yes","","no","english_first","'Book top holiday adventures with SadaPay instantly' - stadium crowd."),
"DcArHXYsPEn":("brand","genz","calendar_topical","national_trust","","pride","none","customer","ugc","no","","no","english_first","Independence-week flag collage reel."),
"DcTcKuIRmbB":("SadaPay customer protection","genz","csr","security_fraud","","humour","none","customer","ugc","no","","yes","other","'Fraud, fake sites aur...' - football-shirt creator piece on scam sites."),
"DQCHneYjOba":("brand","geny","brand_building","national_trust","brand_corporate","pride","claim_statistics","none_people","template_graphic","no","Built on trust","no","english_first","'Built on trust. Grown by you. Swipe to see how Pakistan's favorite digital wallet quietly grew stronger in 2025.' A year-in-review carousel of its own numbers."),
"DT0Qv_CjL8m":("brand","geny","brand_building","brand_corporate","","pride","claim_statistics","none_people","template_graphic","no","Built on trust","no","english_first","'50.4% New Customers Onboarded' - a single performance number as the whole creative."),
"DXPGtwKAX4X":("SadaPay x tapmad","genz","tactical_promo","price_affordability","lifestyle_reward","aspiration","none","none_people","template_graphic","no","","no","english_first","'Get 40% off on the one-month plan - Rs 499 Rs 299. Watch PSL 2026 live on tapmad.'"),
"DZ0SI0Is1XK":("SadaPay debit card","geny","brand_building","brand_corporate","convenience_digital","aspiration","none","none_people","studio","no","","no","english_first","'The card built to go as far as you do' - product film, no offer."),
"DZclfkVNukb":("SadaPay employer brand","genz","recruitment","brand_corporate","","pride","none","none_people","template_graphic","no","","no","english_first","Reposted from @lifeatsadapay - employer-brand content carried on the consumer feed."),
}

posts = list(csv.DictReader(open(SCAN / "posts.csv", encoding="utf-8")))
wb = openpyxl.load_workbook(SCAN / "capture.xlsx")
ws = wb["Ads"]
hdr = [c.value for c in ws[1]]
idx = {h: i + 1 for i, h in enumerate(hdr)}

r = 2
missing = []
for p in posts:
    sc = p["shortcode"]
    if sc not in C:
        missing.append(sc)
        continue
    (pf, gen, ct, c1, c2, reg, proof, frame, prod, eco, plat, obj, lang, notes) = C[sc]
    m = re.search(r"on (\w+ \d{1,2}, \d{4})", p["alt_text"] or "")
    d = _dt.strptime(m.group(1), "%B %d, %Y").strftime("%Y-%m-%d") if m else ""
    rep = p.get("reposted_from") or ""
    if sc == "DZclfkVNukb":
        rep = "lifeatsadapay"
    if sc == "DcaWojblXMd":
        rep = "mercantilepakistan (collab)"
    vals = {"post_id": sc, "brand": p["brand"], "ring": p["ring"], "handle": p["handle"],
            "date": d, "format": "video" if p["kind"] == "reel" else "static",
            "product_focus": pf, "audience_generation": gen, "content_type": ct,
            "claim_primary": c1, "claim_secondary": c2, "register": reg,
            "proof_device": proof, "who_is_in_frame": frame, "production_value": prod,
            "ecosystem_push": eco, "campaign_platform": plat, "addresses_objection": obj,
            "language_lead": lang, "reposted_from": rep, "source_url": p["source_url"],
            "image_file": p["image_file"], "coder": "claude", "notes": notes}
    for k, v in vals.items():
        ws.cell(row=r, column=idx[k], value=v)
    r += 1

wb.save(SCAN / "capture.xlsx")
print("wrote %d coded rows; uncoded: %s" % (r - 2, missing))
