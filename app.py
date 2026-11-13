import os
import re
import csv
import uuid
import zipfile
import traceback
from flask import Flask, render_template, request, send_file, jsonify
import fitz  # PyMuPDF
import openpyxl
import docx

app = Flask(__name__)
BASE_DIR = os.path.dirname(__file__)
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def normalize(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip()

def extract_numeric(val):
    if val is None:
        return None
    num_str = re.sub(r'[^\d]', '', str(val))
    try:
        return int(num_str) if num_str else None
    except ValueError:
        return None

def format_inr(val):
    num = extract_numeric(val)
    if num is not None:
        return f"₹{num:,}"
    s = str(val or '').strip()
    return s if s.startswith("₹") else f"₹{s}" if s else ""

def check_statutory_notice(pdf_text):
    clean = re.sub(r'[^a-zA-Z0-9]', ' ', pdf_text.lower())
    tokens = ["resellers", "free", "set", "price", "below", "mrp"]
    return all(tok in clean for tok in tokens)

def get_model_family(details):
    det = details.lower()
    if "mini" in det:
        return "iPad mini"
    if "ipad pro" in det:
        return "iPad Pro"
    if "air" in det:
        return "iPad Air"
    if "ipad" in det:
        return "iPad (A16)"
    if "pencil" in det:
        return "Apple Pencil"
    if "keyboard" in det:
        return "Keyboards"
    if "folio" in det:
        return "Folio"
    return "Other"

def parse_master_data(excel_files, word_files):
    price_db = []
    model_offers = {}

    for excel_path in excel_files:
        if not excel_path or not os.path.exists(excel_path) or os.path.getsize(excel_path) == 0:
            continue
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                if ws.max_row < 2:
                    continue
                first_row = [cell.value for cell in ws[1] if cell.value is not None]
                headers = {normalize(val).lower(): idx for idx, val in enumerate(first_row)}
                
                def find_idx(keywords):
                    for h, idx in headers.items():
                        for kw in keywords:
                            if kw in h:
                                return idx
                    return None

                mpn_idx = find_idx(["mpn", "part number", "sku"])
                details_idx = find_idx(["product details", "product", "details", "description"])
                mrp_idx = find_idx(["mrp", "price", "regular price"])
                eff_idx = find_idx(["price ticket : cashback effective price on card", "cashback effective price", "effective price"])
                loan_idx = find_idx(["price ticket - loans", "loans", "emi"])
                cb_non_emi = find_idx(["cashback: icici bank credit card non emi", "non emi", "cashback"])
                cb_emi = find_idx(["cashback: icici bank credit card emi", "emi cashback"])
                scheme_idx = find_idx(["schemes", "scheme"])

                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not any(row):
                        continue
                    def get_val(idx):
                        return row[idx] if idx is not None and idx < len(row) else None
                        
                    raw_mrp = get_val(mrp_idx)
                    num_mrp = extract_numeric(raw_mrp)
                    raw_details = normalize(get_val(details_idx))

                    if num_mrp and raw_details:
                        eff_val = extract_numeric(get_val(eff_idx))
                        loan_val = extract_numeric(get_val(loan_idx))
                        
                        cb_val = extract_numeric(get_val(cb_non_emi))
                        if not cb_val:
                            cb_val = extract_numeric(get_val(cb_emi))

                        item = {
                            "mpn": str(get_val(mpn_idx) or "").strip(),
                            "details": raw_details,
                            "mrp_num": num_mrp,
                            "mrp_str": format_inr(num_mrp),
                            "effective_num": eff_val,
                            "loan_num": loan_val,
                            "cashback_num": cb_val
                        }
                        price_db.append(item)

                        family = get_model_family(raw_details)
                        if family not in model_offers:
                            model_offers[family] = {
                                "cashback": cb_val,
                                "loan": loan_val,
                                "effective": eff_val
                            }
                        else:
                            if cb_val and (not model_offers[family]["cashback"] or cb_val > model_offers[family]["cashback"]):
                                model_offers[family]["cashback"] = cb_val
                            if loan_val and (not model_offers[family]["loan"] or loan_val < model_offers[family]["loan"]):
                                model_offers[family]["loan"] = loan_val
                            if eff_val and (not model_offers[family]["effective"] or eff_val < model_offers[family]["effective"]):
                                model_offers[family]["effective"] = eff_val

        except Exception as e:
            print("Notice: Excel parsing notice:", e)

    terms_clauses = []
    for word_path in word_files:
        if not word_path or not os.path.exists(word_path) or os.path.getsize(word_path) == 0:
            continue
        try:
            doc = docx.Document(word_path)
            for p in doc.paragraphs:
                txt = normalize(p.text)
                if txt and len(txt) > 20:
                    terms_clauses.append(txt)
        except Exception:
            pass
        
    return price_db, model_offers, terms_clauses

def is_product_in_pdf(product_details, pdf_text_lower):
    det = product_details.lower()

    if 'pencil' in det:
        if 'pro' in det and 'pencil pro' in pdf_text_lower:
            return True
        if 'usb-c' in det and ('pencil (usb-c)' in pdf_text_lower or 'pencil usb-c' in pdf_text_lower):
            return True
        if ('1st' in det or 'first' in det) and '1st generation' in pdf_text_lower:
            return True
        return 'apple pencil' in pdf_text_lower

    if 'magic keyboard' in det:
        return 'magic keyboard' in pdf_text_lower

    if 'smart folio' in det:
        return 'smart folio' in pdf_text_lower

    if 'bluetooth keyboard' in det or 'keyboards' in det:
        return 'keyboard' in pdf_text_lower

    cap_match = re.search(r'\b(64gb|128gb|256gb|512gb|1tb|2tb)\b', det)
    if cap_match:
        cap = cap_match.group(1)
        if cap not in pdf_text_lower:
            return False

    if '13-inch' in det or '13 inch' in det or '13"' in det:
        if '13-inch' not in pdf_text_lower and '13 inch' not in pdf_text_lower and '13"' not in pdf_text_lower:
            return False

    if 'ipad pro' in det:
        has_ipad_pro = bool(re.search(r'\bipad\s+pro\b(?!\s*model)', pdf_text_lower))
        if not has_ipad_pro:
            return False

    if 'air' in det and 'air' not in pdf_text_lower:
        return False

    if 'mini' in det:
        return 'mini' in pdf_text_lower

    if 'ipad' in det:
        return 'ipad' in pdf_text_lower

    if 'macbook' in det and 'macbook' not in pdf_text_lower:
        return False
    if 'iphone' in det and 'iphone' not in pdf_text_lower:
        return False
    if 'watch' in det and 'watch' not in pdf_text_lower:
        return False

    return True

def audit_single_pdf(pdf_path, original_filename, price_db, model_offers, terms_clauses, session_id):
    safe_base = os.path.splitext(os.path.basename(original_filename))[0]
    safe_base = re.sub(r'[^a-zA-Z0-9_\- ]', '_', safe_base)

    redlined_filename = f"REDLINED_{session_id}_{safe_base}.pdf"
    redlined_path = os.path.join(app.config['OUTPUT_FOLDER'], redlined_filename)

    corrected_filename = f"CORRECTED_{session_id}_{safe_base}.pdf"
    corrected_path = os.path.join(app.config['OUTPUT_FOLDER'], corrected_filename)

    try:
        doc = fitz.open(pdf_path, filetype="pdf")
    except Exception:
        doc = fitz.open(pdf_path)

    if len(doc) == 0:
        return {
            "filename": original_filename,
            "error": "0 pages found in PDF",
            "passed": 0,
            "failed": 1,
            "findings": []
        }

    page = doc[0]
    pdf_full_text = page.get_text("text")
    normalized_full_text = normalize(pdf_full_text)
    pdf_text_lower = normalized_full_text.lower()
    audit_findings = []

    pdf_price_matches = re.findall(r'₹\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?)', pdf_full_text)
    pdf_numeric_prices = set(extract_numeric(p) for p in pdf_price_matches if extract_numeric(p))

    plain_matches = re.findall(r'(?:₹|rs\.?|inr)\s*([0-9,]+)', pdf_full_text, flags=re.IGNORECASE)
    for pm in plain_matches:
        num = extract_numeric(pm)
        if num:
            pdf_numeric_prices.add(num)

    # 1. HARDWARE & ACCESSORY PRICE VERIFICATION
    relevant_products = [
        item for item in price_db 
        if is_product_in_pdf(item["details"], pdf_text_lower)
    ]

    if relevant_products:
        seen_details = set()
        for item in relevant_products:
            simple_name = re.sub(r'\s*-\s*(silver|blue|pink|yellow|midnight|starlight|space grey|space black|black|white).*', '', item["details"], flags=re.IGNORECASE).strip()
            
            key = (simple_name.lower(), item["mrp_num"])
            if key in seen_details:
                continue
            seen_details.add(key)

            exp_mrp = item["mrp_num"]
            matched = (exp_mrp in pdf_numeric_prices)

            audit_findings.append({
                "category": "Product Price Verification",
                "type": simple_name,
                "found": format_inr(exp_mrp) if matched else "Price Mismatch / Not Found in PDF",
                "expected": format_inr(exp_mrp),
                "status": "pass" if matched else "fail"
            })
            
            if not matched:
                for r in page.search_for(simple_name[:15]):
                    annot = page.add_highlight_annot(r)
                    annot.set_colors(stroke=(1, 0, 0))
                    annot.update()
                    note = page.add_text_annot(r.top_right, f"PRICE ERROR: Expected MRP {format_inr(exp_mrp)}")
                    note.update()

    placeholders = page.search_for("₹000.00") + page.search_for("₹000")
    if placeholders:
        for r in placeholders:
            annot = page.add_highlight_annot(r)
            annot.set_colors(stroke=(1, 0, 0))
            annot.update()
        audit_findings.append({
            "category": "Pricing Integrity",
            "type": "Unfilled Placeholders Found",
            "found": f"{len(placeholders)} placeholder(s) '₹000'",
            "expected": "Active approved prices",
            "status": "fail"
        })

    # 2. ACCURATE CASHBACK & PROMOTIONAL AUDIT (CROSS-CHECKED AGAINST EXCEL)
    ticket_devices = []
    if "ipad (a16)" in pdf_text_lower or ("ipad" in pdf_text_lower and "a16" in pdf_text_lower):
        ticket_devices.append("iPad (A16)")
    if "ipad mini" in pdf_text_lower or "mini" in pdf_text_lower:
        ticket_devices.append("iPad mini")

    # A) HEADLINE BANNER CASHBACK CHECK:
    headline_cb_match = re.search(r'up\s+to\s+₹?\s*(\d[\d,]*)\*?\s*instant\s+cashback', normalized_full_text, re.IGNORECASE)
    pdf_headline_cb = extract_numeric(headline_cb_match.group(1)) if headline_cb_match else None

    expected_ticket_cbs = [
        model_offers[dev]["cashback"] 
        for dev in ticket_devices 
        if dev in model_offers and model_offers[dev].get("cashback")
    ]
    expected_max_cb = max(expected_ticket_cbs) if expected_ticket_cbs else None

    if expected_max_cb and pdf_headline_cb:
        cb_matched = (pdf_headline_cb == expected_max_cb)
        audit_findings.append({
            "category": "Affordability / Offers",
            "type": "Headline Instant Cashback Banner",
            "found": f"Up to {format_inr(pdf_headline_cb)} Cashback",
            "expected": f"Up to {format_inr(expected_max_cb)} Cashback (from Excel)",
            "status": "pass" if cb_matched else "fail"
        })
        if not cb_matched and headline_cb_match:
            for r in page.search_for(headline_cb_match.group(0)):
                annot = page.add_highlight_annot(r)
                annot.set_colors(stroke=(1, 0, 0))
                annot.update()
                note = page.add_text_annot(r.top_right, f"CASHBACK MISMATCH: Found {format_inr(pdf_headline_cb)}, expected {format_inr(expected_max_cb)} as per Excel.")
                note.update()

    # B) MODEL-SPECIFIC CASHBACK IN T&C PARAGRAPH:
    cb_sentence_match = re.search(r'instant\s+cashback\s+of\s+₹?\s*(\d[\d,]*)\s*(?:&|and)\s*₹?\s*(\d[\d,]*)\s*is\s+available\s+on\s*([^\.]+)', normalized_full_text, re.IGNORECASE)
    if cb_sentence_match:
        pdf_cb_first = extract_numeric(cb_sentence_match.group(1))
        pdf_cb_second = extract_numeric(cb_sentence_match.group(2))
        
        # Check iPad (A16) cashback
        if "iPad (A16)" in model_offers and model_offers["iPad (A16)"].get("cashback"):
            exp_a16_cb = model_offers["iPad (A16)"]["cashback"]
            a16_cb_match = (pdf_cb_first == exp_a16_cb)
            audit_findings.append({
                "category": "Affordability / Offers",
                "type": "iPad (A16) Specific Cashback in T&C",
                "found": f"{format_inr(pdf_cb_first)} Cashback",
                "expected": f"{format_inr(exp_a16_cb)} Cashback (from Excel)",
                "status": "pass" if a16_cb_match else "fail"
            })
            if not a16_cb_match:
                for r in page.search_for(cb_sentence_match.group(0)[:40]):
                    annot = page.add_highlight_annot(r)
                    annot.set_colors(stroke=(1, 0, 0))
                    annot.update()
                    note = page.add_text_annot(r.top_right, f"T&C CASHBACK ERROR: Stated {format_inr(pdf_cb_first)} for iPad (A16), but Excel specifies {format_inr(exp_a16_cb)}.")
                    note.update()

        # Check iPad mini cashback
        if "iPad mini" in model_offers and model_offers["iPad mini"].get("cashback"):
            exp_mini_cb = model_offers["iPad mini"]["cashback"]
            mini_cb_match = (pdf_cb_second == exp_mini_cb)
            audit_findings.append({
                "category": "Affordability / Offers",
                "type": "iPad mini Specific Cashback in T&C",
                "found": f"{format_inr(pdf_cb_second)} Cashback",
                "expected": f"{format_inr(exp_mini_cb)} Cashback (from Excel)",
                "status": "pass" if mini_cb_match else "fail"
            })

    # C) STARTING EFFECTIVE PRICE:
    starting_match = re.search(r'starting\s*₹?\s*(\d[\d,]*)\*?', normalized_full_text, re.IGNORECASE)
    if starting_match:
        pdf_starting_price = extract_numeric(starting_match.group(1))
        exp_starting_price = None
        if "iPad (A16)" in model_offers and model_offers["iPad (A16)"].get("effective"):
            exp_starting_price = model_offers["iPad (A16)"]["effective"]
        elif "iPad (A16)" in model_offers and model_offers["iPad (A16)"].get("cashback"):
            exp_starting_price = 49900 - model_offers["iPad (A16)"]["cashback"]

        if exp_starting_price and pdf_starting_price:
            starting_matched = (pdf_starting_price == exp_starting_price)
            audit_findings.append({
                "category": "Affordability / Offers",
                "type": "Starting Effective Price Calculation",
                "found": f"Starting {format_inr(pdf_starting_price)}*",
                "expected": f"Starting {format_inr(exp_starting_price)}* (MRP minus {format_inr(model_offers.get('iPad (A16)', {}).get('cashback', 0))} Cashback)",
                "status": "pass" if starting_matched else "fail"
            })

    # D) NO COST EMI:
    emi_match = re.search(r'or\s*₹?\s*(\d[\d,]*)\*?\*?\s*p\.m\.', normalized_full_text, re.IGNORECASE)
    if emi_match:
        pdf_emi = extract_numeric(emi_match.group(1))
        exp_emi = model_offers.get("iPad (A16)", {}).get("loan")
        if exp_emi and pdf_emi:
            emi_matched = (pdf_emi == exp_emi)
            audit_findings.append({
                "category": "Affordability / Offers",
                "type": "Starting Monthly No-Cost EMI",
                "found": f"{format_inr(pdf_emi)} p.m.",
                "expected": f"{format_inr(exp_emi)} p.m. (from Excel)",
                "status": "pass" if emi_matched else "fail"
            })

    # 3. TERMS & CONDITIONS MATCHING
    has_reseller_notice = check_statutory_notice(normalized_full_text)
    statutory_text = "Resellers are free to set their own prices below MRP."
    audit_findings.append({
        "category": "Legal Compliance",
        "type": "Statutory Reseller Notice",
        "found": "Verified in document footer" if has_reseller_notice else "MISSING / UNREADABLE",
        "expected": statutory_text,
        "status": "pass" if has_reseller_notice else "fail"
    })

    has_applecare = "applecare" in pdf_text_lower
    audit_findings.append({
        "category": "Legal Compliance",
        "type": "AppleCare+ Attachment Notice",
        "found": "Verified in footer" if has_applecare else "Missing AppleCare+ prompt",
        "expected": "Want to protect against drops and spills? Ask us about AppleCare+.",
        "status": "pass" if has_applecare else "fail"
    })

    for clause in terms_clauses:
        clause_lower = clause.lower()
        if "cellular" in clause_lower and "cellular" not in pdf_text_lower:
            continue
        if "pencil" in clause_lower and "pencil" not in pdf_text_lower:
            continue
        if "wi-fi 7" in clause_lower and "wi-fi 7" not in pdf_text_lower and "wi‑fi 7" not in pdf_text_lower:
            continue

        clause_words = [w.lower() for w in re.findall(r'\b[a-zA-Z]{4,}\b', clause)[:6]]
        found = all(w in pdf_text_lower for w in clause_words)
        audit_findings.append({
            "category": "Terms & Conditions",
            "type": "Master Clause Compliance",
            "found": "Verified in footnotes" if found else "Missing or altered wording",
            "expected": (clause[:65] + "...") if len(clause) > 65 else clause,
            "status": "pass" if found else "fail"
        })

    doc.save(redlined_path)
    doc.close()

    try:
        doc_fix = fitz.open(pdf_path, filetype="pdf")
    except Exception:
        doc_fix = fitz.open(pdf_path)
    page_fix = doc_fix[0]

    promo_matches_fix = page_fix.search_for("Add Affordability Messaging Here")
    for r in promo_matches_fix:
        page_fix.add_redact_annot(r, fill=(1, 1, 1))

    placeholders_fix = page_fix.search_for("₹000.00") + page_fix.search_for("₹000")
    for r in placeholders_fix:
        page_fix.add_redact_annot(r, fill=(1, 1, 1))

    if promo_matches_fix or placeholders_fix:
        page_fix.apply_redactions()
        for idx, r in enumerate(placeholders_fix):
            rep_price = price_db[idx]["mrp_str"] if idx < len(price_db) else "₹34,900.00"
            page_fix.insert_text(fitz.Point(r.x0, r.y1 - 1.5), rep_price, fontsize=8.5, fontname="helv", color=(0, 0, 0))

    doc_fix.save(corrected_path, garbage=4, deflate=True)
    doc_fix.close()

    failures = sum(1 for f in audit_findings if f['status'] == 'fail')
    passes = len(audit_findings) - failures

    return {
        "filename": original_filename,
        "passed": passes,
        "failed": failures,
        "status": "APPROVED" if failures == 0 else "REJECTED",
        "findings": audit_findings,
        "downloads": {
            "redlined": f"/download/{redlined_filename}",
            "corrected": f"/download/{corrected_filename}"
        },
        "file_paths": {
            "redlined": redlined_path,
            "corrected": corrected_path
        }
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process', methods=['POST'])
def process_ticket():
    try:
        pdf_files = request.files.getlist('pdf')
        excel_files = request.files.getlist('excel')
        word_files = request.files.getlist('word')

        if not pdf_files or len(pdf_files) == 0 or not pdf_files[0].filename:
            return jsonify({"success": False, "error": "At least one PDF file or folder is required."}), 200

        session_id = str(uuid.uuid4())[:8]

        saved_excel_paths = []
        for ef in excel_files:
            if ef and ef.filename:
                path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_{os.path.basename(ef.filename)}")
                ef.save(path)
                saved_excel_paths.append(path)

        saved_word_paths = []
        for wf in word_files:
            if wf and wf.filename:
                path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_{os.path.basename(wf.filename)}")
                wf.save(path)
                saved_word_paths.append(path)

        price_db, model_offers, terms_clauses = parse_master_data(saved_excel_paths, saved_word_paths)

        results = []
        all_redlined_paths = []
        all_corrected_paths = []

        for pf in pdf_files:
            if not pf or not pf.filename:
                continue
            orig_name = os.path.basename(pf.filename)
            if not orig_name.lower().endswith('.pdf') and not pf.mimetype.endswith('pdf'):
                if '.' in orig_name and not orig_name.lower().endswith(('.pdf', '.tmp')):
                    continue
                orig_name += '.pdf'

            local_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_{orig_name}")
            pf.save(local_path)

            file_result = audit_single_pdf(local_path, orig_name, price_db, model_offers, terms_clauses, session_id)
            results.append(file_result)

            if "file_paths" in file_result:
                all_redlined_paths.append(file_result["file_paths"]["redlined"])
                all_corrected_paths.append(file_result["file_paths"]["corrected"])

        if len(results) == 0:
            return jsonify({"success": False, "error": "No valid PDF files found."}), 200

        zip_filename = f"AUDIT_PACKAGE_{session_id}.zip"
        zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for r_path in all_redlined_paths:
                if os.path.exists(r_path):
                    zf.write(r_path, os.path.join("Redlined_Feedback", os.path.basename(r_path)))
            for c_path in all_corrected_paths:
                if os.path.exists(c_path):
                    zf.write(c_path, os.path.join("Corrected_Print_Ready", os.path.basename(c_path)))

        total_files = len(results)
        approved_files = sum(1 for r in results if r.get('status') == 'APPROVED')
        rejected_files = total_files - approved_files

        first_res = results[0] if results else {}
        total_failed = sum(r.get("failed", 0) for r in results)
        total_passed = sum(r.get("passed", 0) for r in results)
        all_findings = []
        for r in results:
            all_findings.extend(r.get("findings", []))

        return jsonify({
            "success": True,
            "summary": {
                "total_checks": total_failed + total_passed,
                "passed": total_passed,
                "failed": total_failed,
                "status": "APPROVED" if total_failed == 0 else "REJECTED"
            },
            "findings": all_findings,
            "downloads": first_res.get("downloads", {
                "redlined": f"/download/{zip_filename}",
                "corrected": f"/download/{zip_filename}"
            }),
            "batch_summary": {
                "total_files": total_files,
                "approved_files": approved_files,
                "rejected_files": rejected_files,
                "zip_download": f"/download/{zip_filename}"
            },
            "files": results
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "success": False,
            "error": f"Audit error: {str(e)}"
        }), 200

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found", 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    print(f"Starting Price Ticket Verification Engine on port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)
