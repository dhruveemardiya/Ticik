import os
import json
import uuid
from pathlib import Path
from django.conf import settings
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
import openpyxl
from .excel_service import ExcelDataManager

excel_manager = ExcelDataManager.get_instance()

@require_http_methods(["GET"])
def customer_status(request):
    """
    GET /api/customers/status/
    Checks if active customer Excel files exist and returns multi-file and multi-sheet status.
    """
    try:
        status = excel_manager.get_status()
        return JsonResponse(status)
    except Exception as e:
        return JsonResponse({"available": False, "error": str(e)}, status=500)

@csrf_exempt
@require_http_methods(["POST"])
def customer_upload(request):
    """
    POST /api/customers/upload/
    Uploads up to 3 Excel files to temporary storage.
    Parses ALL sheets across ALL files, calculates records, and returns full preview.
    DOES NOT modify active storage or delete any existing data.
    """
    # Accept multiple files from request.FILES
    uploaded_files = request.FILES.getlist('files')
    if not uploaded_files:
        # Check individual slot names: file_1, file_2, file_3, or file
        for key in ['file_1', 'file_2', 'file_3', 'file']:
            if key in request.FILES:
                uploaded_files.append(request.FILES[key])

    if not uploaded_files:
        return JsonResponse({"error": "No file uploaded. Please select Excel workbook(s)."}, status=400)

    # Validate file extensions
    for f in uploaded_files:
        if not (f.name.endswith('.xlsx') or f.name.endswith('.xls')):
            return JsonResponse({"error": f"Invalid file '{f.name}'. Only .xlsx or .xls files are supported."}, status=400)

    temp_id = uuid.uuid4().hex[:8]
    files_data = []
    total_physical_rows = 0

    try:
        for idx, f in enumerate(uploaded_files):
            safe_name = f"upload_{temp_id}_{idx+1}_{f.name}"
            temp_path = excel_manager.temp_dir / safe_name

            with open(temp_path, 'wb+') as destination:
                for chunk in f.chunks():
                    destination.write(chunk)

            parsed = excel_manager.parse_full_workbook(temp_path, original_filename=f.name)
            parsed["file_id"] = f"file_{idx+1}"
            parsed["temp_filename"] = safe_name
            files_data.append(parsed)

            total_physical_rows += parsed["total_records"]

        # Calculate true unique merged customer names across all uploaded workbooks using Smart Customer Grouping
        candidate_rows = []
        for fd in files_data:
            t_path = excel_manager.temp_dir / fd["temp_filename"]
            t_wb = openpyxl.load_workbook(t_path, read_only=True, data_only=True)
            for s_info in fd["sheets"]:
                s_name = s_info["name"]
                if s_name not in t_wb.sheetnames:
                    continue
                sh = t_wb[s_name]
                h_idx = s_info["header_row_index"]
                n_col = s_info["name_column"]
                headers = s_info["headers"]
                n_idx = headers.index(n_col) if n_col in headers else (1 if len(headers) > 1 else 0)

                age_idx = None
                sex_idx = None
                for c_i, h in enumerate(headers):
                    h_clean = str(h).strip().upper()
                    if age_idx is None and (h_clean == 'AGE' or h_clean.startswith('AGE ')):
                        age_idx = c_i
                    if sex_idx is None and (h_clean in ('SEX', 'GENDER') or h_clean.startswith('SEX ') or h_clean.startswith('GENDER ')):
                        sex_idx = c_i

                for r_i, r_cells in enumerate(sh.iter_rows(values_only=True)):
                    if r_i <= h_idx and h_idx >= 0:
                        continue
                    row_dict = {
                        headers[ci]: ("" if ci >= len(r_cells) or r_cells[ci] is None else str(r_cells[ci]).strip())
                        for ci in range(len(headers))
                    }
                    if excel_manager._is_invalid_or_header_row(row_dict, headers, n_col):
                        continue
                    if n_idx < len(r_cells) and r_cells[n_idx]:
                        val = str(r_cells[n_idx]).strip()
                        age_val = r_cells[age_idx] if age_idx is not None and age_idx < len(r_cells) else None
                        sex_val = r_cells[sex_idx] if sex_idx is not None and sex_idx < len(r_cells) else None
                        candidate_rows.append({"name": val, "age": age_val, "sex": sex_val})
            t_wb.close()

        smart_grouped = excel_manager.compute_smart_grouped_customers(candidate_rows)
        smart_customer_count = len(smart_grouped)

        # Combine all sheets flat list
        flat_sheets = []
        for fd in files_data:
            for s in fd["sheets"]:
                s_copy = dict(s)
                s_copy["file_name"] = fd["filename"]
                s_copy["file_id"] = fd["file_id"]
                flat_sheets.append(s_copy)

        return JsonResponse({
            "is_multi_file": len(files_data) > 1,
            "total_files": len(files_data),
            "filenames": [fd["filename"] for fd in files_data],
            "filename": files_data[0]["filename"] if len(files_data) == 1 else f"{len(files_data)} Excel Files",
            "sheet_count": sum(fd["sheet_count"] for fd in files_data),
            "customer_count": smart_customer_count,
            "total_records": total_physical_rows,
            "files": files_data,
            "sheets": flat_sheets,
            "temp_filename": files_data[0]["temp_filename"] if len(files_data) == 1 else None
        })

    except Exception as e:
        for fd in files_data:
            p = excel_manager.temp_dir / fd["temp_filename"]
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass
        return JsonResponse({"error": f"Unable to read Excel files: {str(e)}"}, status=400)

@csrf_exempt
@require_http_methods(["POST"])
def customer_save(request):
    """
    POST /api/customers/save/
    Commits 1 to 3 uploaded workbooks permanently.
    """
    try:
        data = json.loads(request.body.decode('utf-8')) if request.body else {}
    except Exception:
        data = request.POST

    files_list = data.get('files')
    if not files_list and data.get('temp_filename'):
        files_list = [{
            "temp_filename": data.get('temp_filename'),
            "filename": data.get('filename', 'customers.xlsx')
        }]

    if not files_list:
        return JsonResponse({"error": "No files provided to save."}, status=400)

    save_items = []
    for item in files_list:
        t_name = item.get('temp_filename')
        f_name = item.get('filename', 'workbook.xlsx')
        t_path = excel_manager.temp_dir / t_name
        if not t_path.exists():
            return JsonResponse({"error": f"Temporary file '{t_name}' not found. Please upload again."}, status=404)
        save_items.append({
            "temp_file_path": t_path,
            "original_filename": f_name
        })

    try:
        metadata = excel_manager.save_multiple_excel_permanently(save_items)
        for item in files_list:
            excel_manager.clean_temp_file(item.get('temp_filename'))

        return JsonResponse({
            "success": True,
            "message": f"{len(save_items)} Excel file(s) permanently saved and indexed.",
            "data": metadata
        })
    except Exception as e:
        return JsonResponse({"error": f"Failed to save customer data: {str(e)}"}, status=500)

@csrf_exempt
@require_http_methods(["POST"])
def customer_cancel_upload(request):
    """
    POST /api/customers/cancel_upload/
    Cancels temporary upload and deletes the temporary files.
    Active data remains untouched.
    """
    try:
        data = json.loads(request.body.decode('utf-8')) if request.body else {}
    except Exception:
        data = request.POST

    temp_files = data.get('temp_filenames') or []
    if data.get('temp_filename'):
        temp_files.append(data.get('temp_filename'))

    if data.get('files') and isinstance(data.get('files'), list):
        for f in data.get('files'):
            if f.get('temp_filename'):
                temp_files.append(f.get('temp_filename'))

    for tf in temp_files:
        excel_manager.clean_temp_file(tf)

    return JsonResponse({"success": True, "message": "Temporary upload cancelled. Active data preserved."})

@require_http_methods(["GET"])
def customer_sheets(request):
    """
    GET /api/customers/sheets/
    Returns list of all sheets across all saved workbooks.
    """
    status = excel_manager.get_status()
    if not status.get("available"):
        return JsonResponse({"available": False, "sheets": []})
    return JsonResponse({
        "available": True,
        "total_files": status.get("total_files", 1),
        "filenames": status.get("filenames", []),
        "sheet_count": status.get("sheet_count", 0),
        "customer_count": status.get("customer_count", 0),
        "total_records": status.get("total_records", 0),
        "sheets": status.get("sheets", [])
    })

@require_http_methods(["GET"])
def customer_list(request):
    """
    GET /api/customers/list/?sheet=<sheet>&search=<query>&limit=<limit>
    Returns unique customer records for left panel and directory modal.
    """
    sheet_name = request.GET.get('sheet', '').strip() or None
    search = request.GET.get('search', '').strip() or None
    limit_raw = request.GET.get('limit')
    limit = int(limit_raw) if limit_raw and limit_raw.isdigit() else None

    try:
        customers = excel_manager.list_customers(sheet_name=sheet_name, search=search, limit=limit)
        return JsonResponse({
            "count": len(customers),
            "sheet_filter": sheet_name,
            "customers": customers
        })
    except Exception as e:
        return JsonResponse({"error": f"Failed to list customers: {str(e)}"}, status=500)

@require_http_methods(["GET"])
def customer_search(request):
    """
    GET /api/customers/search/?name=<name>&sheet=<sheet>
    Searches customers across ALL files and sheets by customer NAME ONLY.
    Returns UNIQUE customer names with record_count.
    """
    query = (request.GET.get('name') or request.GET.get('query') or request.GET.get('q') or '').strip()
    sheet_name = request.GET.get('sheet', '').strip() or None

    if not query:
        return JsonResponse({
            "query": query,
            "sheet_filter": sheet_name,
            "count": 0,
            "results": []
        })

    try:
        search_data = excel_manager.search_customers(query, sheet_name=sheet_name)
        return JsonResponse(search_data)
    except Exception as e:
        return JsonResponse({"error": f"Search failed: {str(e)}"}, status=500)

@require_http_methods(["GET"])
def customer_records_by_name(request):
    """
    GET /api/customers/records/?name=<CustomerName>
    Returns ALL physical Excel records belonging to that unique customer name across all files.
    """
    name = request.GET.get('name', '').strip()
    if not name:
        return JsonResponse({"error": "name parameter is required."}, status=400)

    try:
        data = excel_manager.get_customer_records_by_name(name)
        if not data:
            return JsonResponse({"error": f"Customer '{name}' not found."}, status=404)
        return JsonResponse(data)
    except Exception as e:
        return JsonResponse({"error": f"Failed to fetch customer records: {str(e)}"}, status=500)

@csrf_exempt
@require_http_methods(["GET", "DELETE"])
def customer_record(request):
    """
    GET    /api/customers/record/?record_id=<id>  -> Fetch all columns for exact row
    DELETE /api/customers/record/?record_id=<id>  -> Permanently delete exact physical row from its Excel file
    """
    record_id = request.GET.get('record_id', '').strip()
    if not record_id:
        return JsonResponse({"error": "record_id parameter is required."}, status=400)

    if request.method == "GET":
        try:
            record = excel_manager.get_record(record_id)
            if not record:
                return JsonResponse({"error": f"Record with ID '{record_id}' not found."}, status=404)
            return JsonResponse(record)
        except Exception as e:
            return JsonResponse({"error": f"Failed to fetch record: {str(e)}"}, status=500)

    elif request.method == "DELETE":
        try:
            res = excel_manager.delete_record(record_id)
            return JsonResponse({
                "success": True,
                "message": f"Customer record '{record_id}' permanently deleted from Excel.",
                "data": res
            })
        except KeyError as e:
            return JsonResponse({"error": str(e)}, status=404)
        except Exception as e:
            return JsonResponse({"error": f"Failed to delete customer record: {str(e)}"}, status=500)

@csrf_exempt
@require_http_methods(["POST"])
def customer_replace(request):
    """
    POST /api/customers/replace/
    Allows replacement of customer workbooks.
    """
    try:
        data = json.loads(request.body.decode('utf-8')) if request.body else {}
    except Exception:
        data = request.POST

    if data.get('files') or data.get('temp_filename'):
        return customer_save(request)

    return JsonResponse({
        "success": True,
        "message": "Ready to upload replacement Excel files. Current data is preserved until replacement is confirmed."
    })

@csrf_exempt
@require_http_methods(["GET", "DELETE", "POST"])
def delete_customer_view(request):
    """
    DELETE /api/customers/customer/?name=<Customer Name>
    POST   /api/customers/delete_customer/
    Permanently deletes all physical Excel records belonging to a customer across ALL files and sheets.
    """
    customer_name = request.GET.get('name') or request.POST.get('name') or ''
    if not customer_name and request.body:
        try:
            body = json.loads(request.body.decode('utf-8'))
            customer_name = body.get('name', '')
        except Exception:
            pass
    customer_name = customer_name.strip()
    if not customer_name:
        return JsonResponse({"error": "name parameter is required."}, status=400)

    try:
        res = excel_manager.delete_customer(customer_name)
        return JsonResponse({
            "success": True,
            "message": f"All records for customer '{customer_name}' permanently deleted from Excel.",
            "data": res
        })
    except KeyError as e:
        return JsonResponse({"error": str(e)}, status=404)
    except Exception as e:
        return JsonResponse({"error": f"Failed to delete customer: {str(e)}"}, status=500)

@csrf_exempt
@require_http_methods(["GET", "DELETE", "POST"])
def delete_all_customers_view(request):
    """
    DELETE /api/customers/all/
    POST   /api/customers/delete_all/
    Permanently deletes ALL customer records across ALL sheets in all Excel workbooks.
    """
    try:
        res = excel_manager.delete_all_customers()
        return JsonResponse({
            "success": True,
            "message": "All customer records permanently deleted from Excel.",
            "data": res
        })
    except Exception as e:
        return JsonResponse({"error": f"Failed to delete all customers: {str(e)}"}, status=500)
