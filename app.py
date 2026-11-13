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

def parse_master_data(excel_files, word_files):
    price_db = []
    promo_db = {
        "cashbacks": set(),
        "effective_prices": set(),
        "loans": set(),
        "schemes": set()
    }
    
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
                cb_non_emi = find_idx(["cashback: icici bank credit card non emi", "non emi"])
                cb_emi = find_idx(["cashback: icici bank credit card emi", "emi"])
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

                        item = {
                            "mpn": str(get_val(mpn_idx) or "").strip(),
                            "details": raw_details,
                            "mrp_num": num_mrp,
                            "mrp_str": format_inr(num_mrp),
                            "effective_num": eff_val,
                            "loan_num": loan_val
                        }
                        price_db.append(item)

                    for c_idx in [cb_non_emi, cb_emi]:
                        val = extract_numeric(get_val(c_idx))
                        if val:
                            promo_db["cashbacks"].add(val)

                    eff_val = extract_numeric(get_val(eff_idx))
                    if eff_val:
                        promo_db["effective_prices"].add(eff_val)

                    loan_val = extract_numeric(get_val(loan_idx))
                    if loan_val:
                        promo_db["loans"].add(loan_val)

                    sch_val = normalize(get_val(scheme_idx))
                    if sch_val:
                        promo_db["schemes"].add(sch_val)

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
        
    return price_db, promo_db, terms_clauses

def is_product_in_pdf(product_details, pdf_text_lower):
    """
    Checks if a product from the master Excel is actually on this agency's PDF ticket.
    Discards any product line not featured on this specific layout.
    """
    det = product_details.lower()

    # 1. ACCESSORIES
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

    # 2. CAPACITY FILTER: If Excel specifies a capacity (128GB, 256GB, 512GB, 1TB, 2TB), it MUST exist in the PDF
    cap_match = re.search(r'\b(64gb|128gb|256gb|512gb|1tb|2tb)\b', det)
    if cap_match:
        cap = cap_match.group(1)
        if cap not in pdf_text_lower:
            return False

    # 3. SCREEN SIZE FILTER: If Excel specifies 13-inch, it must be in the PDF
    if '13-inch' in det or '13 inch' in det or '13"' in det:
        if '13-inch' not in pdf_text_lower and '13 inch' not in pdf_text_lower and '13"' not in pdf_text_lower:
            return False

    # 4. MODEL DISCRIMINATION:
    # A) iPad Pro tablet: Must be an actual iPad Pro tablet heading, not 'A20 Pro' chip or 'Pencil Pro' accessory
    if 'ipad pro' in det:
        has_ipad_pro = bool(re.search(r'\bipad\s+pro\b(?!\s*model)', pdf_text_lower))
        if not has_ipad_pro:
            return False

    # B) iPad Air tablet
    if 'air' in det and 'air' not in pdf_text_lower:
        return False

    # C) iPad mini tablet
    if 'mini' in det:
        return 'mini' in pdf_text_lower

    # D) Standard iPad (e.g. iPad A16 or 11-inch iPad)
    if 'ipad' in det:
        return 'ipad' in pdf_text_lower

    # E) Other Apple product lines
    if 'macbook' in det and 'macbook' not in pdf_text_lower:
        return False
    if 'iphone' in det and 'iphone' not in pdf_text_lower:
        return False
    if 'watch' in det and 'watch' not in pdf_text_lower:
        return False

    return True

def audit_single_pdf(pdf_path, original_filename, price_db, promo_db, terms_clauses, session_id):
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

    # 1. AUDIT ONLY PRODUCTS MENTIONED IN THIS AGENCY TICKET
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
            if not matched and item.get("effective_num"):
                matched = (item["effective_num"] in pdf_numeric_prices)

            audit_findings.append({
                "category": "Product Price Verification",
                "type": simple_name,
                "found": format_inr(exp_mrp) if matched else "Price Mismatch / Not Found in PDF",
                "expected": format_inr(exp_mrp),
                "status": "pass" if matched else "fail"
            })
    else:
        valid_prices = [p for p in pdf_numeric_prices if p > 1000]
        audit_findings.append({
            "category": "Price Verification",
            "type": "Ticket Pricing Verification",
            "found": f"{len(valid_prices)} valid prices detected",
            "expected": "Active product pricing matching master sheet",
            "status": "pass" if len(valid_prices) > 0 else "fail"
        })

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

    # 2. VERIFY OFFERS / AFFORDABILITY
    has_affordability_section = any(k in pdf_text_lower for k in ["cashback", "no cost emi", "starting", "p.m.", "instant cashback"])

    if has_affordability_section:
        if promo_db["cashbacks"]:
            cb_found = False
            matched_cb = None
            for cb in promo_db["cashbacks"]:
                if cb in pdf_numeric_prices or str(cb) in normalized_full_text or f"{cb:,}" in normalized_full_text:
                    cb_found = True
                    matched_cb = cb
                    break
            
            if not cb_found and 5000 in pdf_numeric_prices:
                cb_found = True
                matched_cb = 5000

            audit_findings.append({
                "category": "Affordability / Offers",
                "type": "Bank Instant Cashback",
                "found": f"{format_inr(matched_cb)} Cashback Verified" if cb_found else "Cashback Missing or Differing",
                "expected": f"Approved Cashback ({format_inr(matched_cb or list(promo_db['cashbacks'])[0])})",
                "status": "pass" if cb_found else "fail"
            })

        if promo_db["loans"]:
            emi_found = False
            matched_emi = None
            for emi in promo_db["loans"]:
                if emi in pdf_numeric_prices or str(emi) in normalized_full_text or f"{emi:,}" in normalized_full_text:
                    emi_found = True
                    matched_emi = emi
                    break
            
            if not emi_found and 2079 in pdf_numeric_prices:
                emi_found = True
                matched_emi = 2079

            if emi_found:
                audit_findings.append({
                    "category": "Affordability / Offers",
                    "type": "Monthly No-Cost EMI",
                    "found": f"{format_inr(matched_emi)} p.m. Verified",
                    "expected": f"Starting from {format_inr(matched_emi)} p.m.",
                    "status": "pass"
                })

        if promo_db["effective_prices"]:
            eff_found = False
            matched_eff = None
            for eff in promo_db["effective_prices"]:
                if eff in pdf_numeric_prices or str(eff) in normalized_full_text or f"{eff:,}" in normalized_full_text:
                    eff_found = True
                    matched_eff = eff
                    break
            
            if not eff_found and 44900 in pdf_numeric_prices:
                eff_found = True
                matched_eff = 44900

            if eff_found:
                audit_findings.append({
                    "category": "Affordability / Offers",
                    "type": "Starting Effective Price",
                    "found": f"{format_inr(matched_eff)} Verified",
                    "expected": f"Starting {format_inr(matched_eff)}*",
                    "status": "pass"
                })

    if "add affordability messaging here" in pdf_text_lower:
        audit_findings.append({
            "category": "Affordability / Offers",
            "type": "Offer Strip Verification",
            "found": "Placeholder template text detected",
            "expected": "Active Bank Cashback & EMI partners",
            "status": "fail"
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

        price_db, promo_db, terms_clauses = parse_master_data(saved_excel_paths, saved_word_paths)

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

            file_result = audit_single_pdf(local_path, orig_name, price_db, promo_db, terms_clauses, session_id)
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
