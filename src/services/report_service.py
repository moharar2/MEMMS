"""
Enterprise Report & Analytics Service.
Generates comprehensive reports, statistical aggregations, and exports (Excel, CSV, PDF/HTML).
Fully localized, explicit column mapping, zero emojis.
"""
import io
import csv
from datetime import datetime, date, timedelta
from typing import List, Dict, Any, Optional
from sqlalchemy.orm import Session
from sqlalchemy import func, or_, and_

from src.models.equipment import Equipment, EquipmentTransfer
from src.models.work_order import WorkOrder, WorkOrderPart
from src.models.preventive import PMPlan, PMSchedule, PMExecution
from src.models.inventory import SparePart, InventoryTransaction
from src.models.hospital import Department, Hospital
from src.models.technician import Technician
from src.models.vendor import Vendor
from src.core.i18n import tr


class ReportService:
    # ── 1. Equipment Inventory Report ────────────────────────────────────────
    def get_equipment_report_data(
        self,
        db: Session,
        department_id: Optional[int] = None,
        status: Optional[str] = None,
        risk_level: Optional[str] = None,
        category: Optional[str] = None
    ) -> Dict[str, Any]:
        query = db.query(Equipment)
        if department_id:
            query = query.filter(Equipment.department_id == department_id)
        if status and status != "all":
            query = query.filter(Equipment.status == status)
        if risk_level and risk_level != "all":
            query = query.filter(Equipment.risk_level == risk_level)
        if category and category != "all":
            query = query.filter(Equipment.category == category)

        items = query.order_by(Equipment.id.desc()).all()

        total_count = len(items)
        operational_count = sum(1 for e in items if e.status == "operational")
        in_maintenance = sum(1 for e in items if e.status == "under_maintenance")
        out_of_service = sum(1 for e in items if e.status == "out_of_service")
        total_value = sum(float(e.purchase_price or 0.0) for e in items)

        column_keys = [
            "asset_number", "name_ar", "model", "serial_number", "department",
            "category", "status", "risk_level", "purchase_price", "warranty_expiry"
        ]

        rows = []
        for eq in items:
            dept_name = eq.department.name_ar if eq.department else "-"
            status_text = tr(f"status.{eq.status}", default=eq.status)
            risk_text = tr(f"risk.{eq.risk_level}", default=eq.risk_level)

            rows.append({
                "id": eq.id,
                "asset_number": eq.asset_number or "-",
                "name_ar": eq.name_ar or "-",
                "model": eq.model or "-",
                "serial_number": eq.serial_number or "-",
                "department": dept_name,
                "category": eq.category or "-",
                "status": status_text,
                "risk_level": risk_text,
                "purchase_price": float(eq.purchase_price or 0.0),
                "purchase_date": eq.purchase_date.strftime("%Y-%m-%d") if eq.purchase_date else "-",
                "warranty_expiry": eq.warranty_expiry_date.strftime("%Y-%m-%d") if eq.warranty_expiry_date else "-",
            })

        return {
            "title": tr("reports.equipment_inventory_report"),
            "column_keys": column_keys,
            "summary": {
                "total_count": total_count,
                "operational_count": operational_count,
                "in_maintenance": in_maintenance,
                "out_of_service": out_of_service,
                "total_value": total_value,
                "display": {
                    "إجمالي الأجهزة": total_count,
                    "الأجهزة العاملة": operational_count,
                    "تحت الصيانة": in_maintenance,
                    "خارج الخدمة": out_of_service,
                    "القيمة الإجمالية للأصول": f"{total_value:,.2f} ر.س"
                }
            },
            "headers": [
                tr("equipment.asset_id"),
                tr("equipment.name_ar"),
                tr("equipment.model"),
                tr("equipment.serial_number"),
                tr("equipment.department"),
                tr("equipment.category"),
                tr("equipment.status"),
                tr("equipment.risk_level"),
                tr("equipment.purchase_price"),
                tr("equipment.warranty_expiry")
            ],
            "rows": rows
        }

    # ── 2. Work Orders Report ────────────────────────────────────────────────
    def get_work_orders_report_data(
        self,
        db: Session,
        from_date: Optional[date] = None,
        to_date: Optional[date] = None,
        status: Optional[str] = None,
        priority: Optional[str] = None,
        department_id: Optional[int] = None
    ) -> Dict[str, Any]:
        query = db.query(WorkOrder)

        if from_date:
            query = query.filter(WorkOrder.created_at >= datetime.combine(from_date, datetime.min.time()))
        if to_date:
            query = query.filter(WorkOrder.created_at <= datetime.combine(to_date, datetime.max.time()))
        if status and status != "all":
            query = query.filter(WorkOrder.status == status)
        if priority and priority != "all":
            query = query.filter(WorkOrder.priority == priority)
        if department_id:
            query = query.join(Equipment).filter(Equipment.department_id == department_id)

        orders = query.order_by(WorkOrder.id.desc()).all()

        total_orders = len(orders)
        completed_count = sum(1 for o in orders if o.status in ("completed", "closed"))
        open_count = sum(1 for o in orders if o.status in ("open", "pending", "assigned"))
        in_progress = sum(1 for o in orders if o.status == "in_progress")
        total_cost = sum(float(o.total_cost or 0.0) for o in orders)
        parts_cost = sum(float(o.total_parts_cost or 0.0) for o in orders)
        labor_cost = sum(float(o.total_labor_cost or 0.0) for o in orders)

        column_keys = [
            "work_order_number", "equipment", "priority", "status", "reported_date",
            "technician", "failure_description", "completed_date", "parts_cost",
            "labor_cost", "total_cost"
        ]

        rows = []
        for o in orders:
            eq_name = f"{o.equipment.name_ar} ({o.equipment.asset_number})" if o.equipment else "-"
            tech_name = o.assigned_technician.full_name if o.assigned_technician else "-"
            created_str = o.created_at.strftime("%Y-%m-%d") if o.created_at else "-"
            completed_str = o.completed_at.strftime("%Y-%m-%d") if o.completed_at else "-"
            status_text = tr(f"status.{o.status}", default=o.status)
            priority_text = tr(f"priority.{o.priority}", default=o.priority)

            rows.append({
                "id": o.id,
                "work_order_number": o.work_order_number,
                "equipment": eq_name,
                "priority": priority_text,
                "status": status_text,
                "reported_date": created_str,
                "technician": tech_name,
                "failure_description": o.failure_description or "-",
                "completed_date": completed_str,
                "parts_cost": float(o.total_parts_cost or 0.0),
                "labor_cost": float(o.total_labor_cost or 0.0),
                "total_cost": float(o.total_cost or 0.0)
            })

        return {
            "title": tr("reports.work_orders_summary_report"),
            "column_keys": column_keys,
            "summary": {
                "total_orders": total_orders,
                "completed_count": completed_count,
                "open_count": open_count,
                "in_progress": in_progress,
                "parts_cost": parts_cost,
                "labor_cost": labor_cost,
                "total_cost": total_cost,
                "display": {
                    "إجمالي أوامر العمل": total_orders,
                    "الأوامر المكتملة": completed_count,
                    "الأوامر المفتوحة": open_count,
                    "قيد التنفيذ": in_progress,
                    "تكلفة قطع الغيار": f"{parts_cost:,.2f} ر.س",
                    "تكلفة العمالة": f"{labor_cost:,.2f} ر.س",
                    "التكلفة الإجمالية للصيانة": f"{total_cost:,.2f} ر.س"
                }
            },
            "headers": [
                tr("work_orders.order_number"),
                tr("work_orders.equipment"),
                tr("common.priority"),
                tr("common.status"),
                tr("work_orders.report_date"),
                tr("work_orders.assigned_to"),
                tr("work_orders.problem_description"),
                tr("work_orders.completed_at"),
                tr("work_orders.parts_cost"),
                tr("work_orders.labor_cost"),
                tr("work_orders.total_cost")
            ],
            "rows": rows
        }

    # ── 3. Preventive Maintenance Compliance Report ──────────────────────────
    def get_pm_compliance_report_data(
        self,
        db: Session,
        from_date: Optional[date] = None,
        to_date: Optional[date] = None,
        status: Optional[str] = None
    ) -> Dict[str, Any]:
        query = db.query(PMSchedule)

        if from_date:
            query = query.filter(PMSchedule.due_date >= from_date)
        if to_date:
            query = query.filter(PMSchedule.due_date <= to_date)
        if status and status != "all":
            query = query.filter(PMSchedule.status == status)

        schedules = query.order_by(PMSchedule.due_date.desc()).all()

        total_pm = len(schedules)
        completed_pm = sum(1 for p in schedules if p.status == "completed")
        pending_pm = sum(1 for p in schedules if p.status == "scheduled")
        overdue_pm = sum(1 for p in schedules if p.status == "overdue")
        compliance_rate = round((completed_pm / total_pm * 100), 1) if total_pm > 0 else 100.0

        column_keys = [
            "scheduled_date", "equipment", "plan_title", "status",
            "result", "technician", "executed_date", "notes"
        ]

        rows = []
        for s in schedules:
            plan_title = s.plan.title if s.plan else "-"
            eq_name = f"{s.equipment.name_ar} ({s.equipment.asset_number})" if s.equipment else "-"
            sched_str = s.due_date.strftime("%Y-%m-%d") if s.due_date else "-"
            status_text = tr(f"status.{s.status}", default=s.status)

            tech_name = "-"
            exec_str = "-"
            result_str = "-"
            notes_str = "-"
            if s.execution:
                tech_name = s.execution.technician.full_name if s.execution.technician else "-"
                exec_str = s.execution.execution_date.strftime("%Y-%m-%d") if s.execution.execution_date else "-"
                result_str = tr("pm.passed") if s.execution.passed else tr("pm.failed")
                notes_str = s.execution.results_summary or s.execution.notes or "-"

            rows.append({
                "id": s.id,
                "scheduled_date": sched_str,
                "equipment": eq_name,
                "plan_title": plan_title,
                "status": status_text,
                "result": result_str,
                "technician": tech_name,
                "executed_date": exec_str,
                "notes": notes_str
            })

        return {
            "title": tr("reports.pm_compliance_report"),
            "column_keys": column_keys,
            "summary": {
                "total_scheduled": total_pm,
                "completed_count": completed_pm,
                "pending_count": pending_pm,
                "overdue_count": overdue_pm,
                "compliance_rate": compliance_rate,
                "display": {
                    "إجمالي المهام المجدولة": total_pm,
                    "المهام المنجزة": completed_pm,
                    "المهام المعلقة": pending_pm,
                    "المهام المتأخرة": overdue_pm,
                    "نسبة الالتزام بالصيانة الوقائية": f"{compliance_rate}%"
                }
            },
            "headers": [
                tr("pm.scheduled_date"),
                tr("pm.equipment"),
                tr("pm.plan_title"),
                tr("common.status"),
                tr("pm.result"),
                tr("pm.performed_by"),
                tr("pm.execution_date"),
                tr("common.notes")
            ],
            "rows": rows
        }

    # ── 4. Inventory & Spare Parts Valuation Report ──────────────────────────
    def get_inventory_report_data(
        self,
        db: Session,
        category: Optional[str] = None,
        low_stock_only: bool = False
    ) -> Dict[str, Any]:
        query = db.query(SparePart).filter(SparePart.is_active.is_(True))

        if category and category != "all":
            query = query.filter(SparePart.category == category)
        if low_stock_only:
            query = query.filter(SparePart.quantity_on_hand <= SparePart.minimum_reorder_level)

        parts = query.order_by(SparePart.id.desc()).all()

        total_parts = len(parts)
        low_stock_count = sum(1 for p in parts if (p.quantity_on_hand or 0) <= (p.minimum_reorder_level or 0))
        total_valuation = sum(float(p.quantity_on_hand or 0) * float(p.unit_cost or 0.0) for p in parts)
        total_quantity = sum(float(p.quantity_on_hand or 0) for p in parts)

        column_keys = [
            "part_number", "name_ar", "category", "quantity", "min_quantity",
            "unit_price", "total_valuation", "location", "vendor"
        ]

        rows = []
        for p in parts:
            stock_qty = p.quantity_on_hand or 0
            unit_price = float(p.unit_cost or 0.0)
            total_cost = round(stock_qty * unit_price, 2)
            vendor_name = p.manufacturer.name if p.manufacturer else "-"

            rows.append({
                "id": p.id,
                "part_number": p.part_number or "-",
                "name_ar": p.name_ar or "-",
                "name_en": p.name_en or "-",
                "category": p.category or "-",
                "quantity": f"{stock_qty:g} {p.unit_of_measure or ''}",
                "min_quantity": f"{p.minimum_reorder_level or 0:g} {p.unit_of_measure or ''}",
                "unit_price": unit_price,
                "total_valuation": total_cost,
                "location": p.storage_bin_location or "-",
                "vendor": vendor_name,
                "is_low_stock": stock_qty <= (p.minimum_reorder_level or 0)
            })

        return {
            "title": tr("reports.inventory_valuation_report"),
            "column_keys": column_keys,
            "summary": {
                "total_parts": total_parts,
                "low_stock_count": low_stock_count,
                "total_quantity": total_quantity,
                "total_valuation": total_valuation,
                "display": {
                    "إجمالي أصناف قطع الغيار": total_parts,
                    "الأصناف منخفضة المخزون": low_stock_count,
                    "إجمالي الكميات بالمخزن": total_quantity,
                    "القيمة الإجمالية للمخزون": f"{total_valuation:,.2f} ر.س"
                }
            },
            "headers": [
                tr("inventory.part_number"),
                tr("inventory.name_ar"),
                tr("inventory.category"),
                tr("inventory.quantity"),
                tr("inventory.reorder_level"),
                tr("inventory.unit_cost"),
                tr("inventory.total_value"),
                tr("inventory.bin_location"),
                tr("inventory.manufacturer")
            ],
            "rows": rows
        }

    # ── 5. Manufacturer Report ───────────────────────────────────────────────
    def get_manufacturer_report_data(
        self,
        db: Session,
        manufacturer_id: Optional[int] = None,
        department_id: Optional[int] = None
    ) -> Dict[str, Any]:
        query = db.query(Equipment)
        if manufacturer_id and manufacturer_id != "all":
            query = query.filter(Equipment.manufacturer_id == manufacturer_id)
        if department_id:
            query = query.filter(Equipment.department_id == department_id)

        items = query.order_by(Equipment.manufacturer_id.asc(), Equipment.id.desc()).all()
        today = date.today()

        total_count = len(items)
        operational_count = sum(1 for e in items if e.status == "operational")
        under_warranty = sum(1 for e in items if e.warranty_expiry_date and e.warranty_expiry_date >= today)
        total_value = sum(float(e.purchase_price or 0.0) for e in items)

        column_keys = [
            "manufacturer", "asset_number", "name_ar", "model", "serial_number",
            "department", "warranty_status", "status", "purchase_price"
        ]

        rows = []
        for eq in items:
            manuf_name = eq.manufacturer.name if eq.manufacturer else "غير محدد"
            dept_name = eq.department.name_ar if eq.department else "-"
            w_status = "ساري الضمان" if (eq.warranty_expiry_date and eq.warranty_expiry_date >= today) else "منتهي الضمان"
            if not eq.warranty_expiry_date:
                w_status = "غير محدد"

            rows.append({
                "id": eq.id,
                "manufacturer": manuf_name,
                "asset_number": eq.asset_number or "-",
                "name_ar": eq.name_ar or "-",
                "model": eq.model or "-",
                "serial_number": eq.serial_number or "-",
                "department": dept_name,
                "warranty_status": w_status,
                "status": tr(f"status.{eq.status}", default=eq.status),
                "purchase_price": float(eq.purchase_price or 0.0)
            })

        return {
            "title": "تقرير الأجهزة الطبية حسب الشركة المصنعة",
            "column_keys": column_keys,
            "summary": {
                "total_count": total_count,
                "operational_count": operational_count,
                "under_warranty": under_warranty,
                "total_value": total_value,
                "display": {
                    "إجمالي الأجهزة": total_count,
                    "الأجهزة العاملة": operational_count,
                    "تحت الضمان": under_warranty,
                    "إجمالي القيمة التقديرية": f"{total_value:,.2f} ر.س"
                }
            },
            "headers": [
                "الشركة المصنعة",
                "رقم الأصل",
                "اسم الجهاز",
                "الموديل",
                "الرقم التسلسلي",
                "القسم",
                "حالة الضمان",
                "الحالة التشغيلية",
                "سعر الشراء"
            ],
            "rows": rows
        }

    # ── 6. Vendor & Agent Report ─────────────────────────────────────────────
    def get_vendor_agent_report_data(
        self,
        db: Session,
        vendor_id: Optional[int] = None,
        from_date: Optional[date] = None,
        to_date: Optional[date] = None
    ) -> Dict[str, Any]:
        query = db.query(Equipment)
        if vendor_id and vendor_id != "all":
            query = query.filter(or_(Equipment.agent_id == vendor_id, Equipment.supplier_id == vendor_id))

        items = query.order_by(Equipment.id.desc()).all()

        total_count = len(items)
        operational_count = sum(1 for e in items if e.status == "operational")
        in_maintenance = sum(1 for e in items if e.status == "under_maintenance")

        column_keys = [
            "vendor", "vendor_type", "asset_number", "name_ar", "model",
            "engineer_name", "engineer_phone", "status", "wo_count"
        ]

        rows = []
        for eq in items:
            agent = eq.agent or eq.supplier
            vendor_name = agent.name if agent else "غير محدد"
            v_type = tr(f"vendor.{agent.vendor_type}", default=agent.vendor_type) if agent else "-"
            eng_name = eq.assigned_engineer.name if eq.assigned_engineer else (agent.contact_person if agent else "-")
            eng_phone = eq.assigned_engineer.phone if eq.assigned_engineer else (agent.phone if agent else "-")
            wo_count = len(eq.work_orders) if eq.work_orders else 0

            rows.append({
                "id": eq.id,
                "vendor": vendor_name,
                "vendor_type": v_type,
                "asset_number": eq.asset_number or "-",
                "name_ar": eq.name_ar or "-",
                "model": eq.model or "-",
                "engineer_name": eng_name or "-",
                "engineer_phone": eng_phone or "-",
                "status": tr(f"status.{eq.status}", default=eq.status),
                "wo_count": wo_count
            })

        return {
            "title": "تقرير الأجهزة المرتبطة بالوكلاء والموردين",
            "column_keys": column_keys,
            "summary": {
                "total_count": total_count,
                "operational_count": operational_count,
                "in_maintenance": in_maintenance,
                "display": {
                    "إجمالي الأجهزة المسندة": total_count,
                    "الأجهزة العاملة": operational_count,
                    "تحت الصيانة": in_maintenance
                }
            },
            "headers": [
                "الوكيل / المورد",
                "نوع النشاط",
                "رقم الأصل",
                "اسم الجهاز",
                "الموديل",
                "مهندس الدعم المسؤول",
                "هاتف التواصل",
                "الحالة التشغيلية",
                "أوامر العمل المسجلة"
            ],
            "rows": rows
        }

    # ── 7. Category / Type Report ────────────────────────────────────────────
    def get_category_report_data(
        self,
        db: Session,
        category: Optional[str] = None,
        department_id: Optional[int] = None
    ) -> Dict[str, Any]:
        query = db.query(Equipment)
        if category and category != "all":
            query = query.filter(Equipment.category == category)
        if department_id:
            query = query.filter(Equipment.department_id == department_id)

        items = query.order_by(Equipment.category.asc(), Equipment.id.desc()).all()

        total_count = len(items)
        operational_count = sum(1 for e in items if e.status == "operational")
        high_risk_count = sum(1 for e in items if e.risk_level == "high")

        column_keys = [
            "category", "asset_number", "name_ar", "manufacturer",
            "model", "department", "status", "risk_level"
        ]

        rows = []
        for eq in items:
            dept_name = eq.department.name_ar if eq.department else "-"
            manuf_name = eq.manufacturer.name if eq.manufacturer else "-"

            rows.append({
                "id": eq.id,
                "category": eq.category or "-",
                "asset_number": eq.asset_number or "-",
                "name_ar": eq.name_ar or "-",
                "manufacturer": manuf_name,
                "model": eq.model or "-",
                "department": dept_name,
                "status": tr(f"status.{eq.status}", default=eq.status),
                "risk_level": tr(f"risk.{eq.risk_level}", default=eq.risk_level)
            })

        return {
            "title": "تقرير الأجهزة الطبية حسب نوع الجهاز والتصنيف",
            "column_keys": column_keys,
            "summary": {
                "total_count": total_count,
                "operational_count": operational_count,
                "high_risk_count": high_risk_count,
                "display": {
                    "إجمالي الأجهزة": total_count,
                    "الأجهزة العاملة": operational_count,
                    "أجهزة عالية الخطورة": high_risk_count
                }
            },
            "headers": [
                "نوع الجهاز / التصنيف",
                "رقم الأصل",
                "اسم الجهاز",
                "الشركة المصنعة",
                "الموديل",
                "القسم",
                "الحالة التشغيلية",
                "مستوى الخطورة"
            ],
            "rows": rows
        }

    # ── 8. Department Summary Report ─────────────────────────────────────────
    def get_department_report_data(
        self,
        db: Session,
        department_id: Optional[int] = None
    ) -> Dict[str, Any]:
        dept_query = db.query(Department)
        if department_id:
            dept_query = dept_query.filter(Department.id == department_id)
        departments = dept_query.order_by(Department.name_ar.asc()).all()

        column_keys = [
            "department", "code", "total_equipment", "operational_count",
            "down_count", "open_work_orders", "overdue_pm"
        ]

        today = date.today()
        rows = []
        grand_total = 0
        grand_operational = 0
        grand_down = 0
        grand_open_wo = 0

        for dept in departments:
            eq_list = dept.equipment or []
            total_eq = len(eq_list)
            op_count = sum(1 for e in eq_list if e.status == "operational")
            down_count = sum(1 for e in eq_list if e.status in ("out_of_service", "under_maintenance"))

            eq_ids = [e.id for e in eq_list]
            open_wo = 0
            overdue_pm = 0
            if eq_ids:
                open_wo = db.query(WorkOrder).filter(
                    WorkOrder.equipment_id.in_(eq_ids),
                    WorkOrder.status.in_(("open", "in_progress", "pending", "assigned"))
                ).count()
                overdue_pm = db.query(PMSchedule).filter(
                    PMSchedule.equipment_id.in_(eq_ids),
                    PMSchedule.status == "overdue"
                ).count()

            grand_total += total_eq
            grand_operational += op_count
            grand_down += down_count
            grand_open_wo += open_wo

            rows.append({
                "id": dept.id,
                "department": dept.name_ar,
                "code": dept.code or "-",
                "total_equipment": total_eq,
                "operational_count": op_count,
                "down_count": down_count,
                "open_work_orders": open_wo,
                "overdue_pm": overdue_pm
            })

        return {
            "title": "تقرير الأجهزة والصيانة الشامل حسب القسم الطبي",
            "column_keys": column_keys,
            "summary": {
                "departments_count": len(departments),
                "grand_total": grand_total,
                "grand_operational": grand_operational,
                "grand_down": grand_down,
                "grand_open_wo": grand_open_wo,
                "display": {
                    "عدد الأقسام المشمولة": len(departments),
                    "إجمالي الأجهزة بالمستشفى": grand_total,
                    "إجمالي الأجهزة العاملة": grand_operational,
                    "إجمالي الأجهزة المتوقفة": grand_down,
                    "أوامر العمل المفتوحة": grand_open_wo
                }
            },
            "headers": [
                "القسم الطبي",
                "كود القسم",
                "إجمالي الأجهزة",
                "الأجهزة العاملة",
                "الأجهزة المتوقفة",
                "أوامر العمل المفتوحة",
                "صيانة وقائية متأخرة"
            ],
            "rows": rows
        }

    # ── 9. Warranty Status & Expiry Report ───────────────────────────────────
    def get_warranty_report_data(
        self,
        db: Session,
        warranty_filter: str = "all",
        days_threshold: int = 60,
        department_id: Optional[int] = None
    ) -> Dict[str, Any]:
        today = date.today()
        query = db.query(Equipment)
        if department_id:
            query = query.filter(Equipment.department_id == department_id)

        if warranty_filter == "under_warranty":
            query = query.filter(Equipment.warranty_expiry_date >= today)
        elif warranty_filter == "expired":
            query = query.filter(Equipment.warranty_expiry_date < today)
        elif warranty_filter == "expiring_soon":
            limit_date = today + timedelta(days=days_threshold)
            query = query.filter(Equipment.warranty_expiry_date >= today, Equipment.warranty_expiry_date <= limit_date)

        items = query.order_by(Equipment.warranty_expiry_date.asc()).all()

        total_count = len(items)
        under_warranty_count = sum(1 for e in items if e.warranty_expiry_date and e.warranty_expiry_date >= today)
        expired_count = sum(1 for e in items if e.warranty_expiry_date and e.warranty_expiry_date < today)

        column_keys = [
            "asset_number", "name_ar", "model", "serial_number",
            "department", "vendor", "warranty_expiry", "days_remaining", "warranty_status"
        ]

        rows = []
        for eq in items:
            dept_name = eq.department.name_ar if eq.department else "-"
            agent = eq.agent or eq.supplier
            vendor_name = agent.name if agent else "-"
            w_date_str = eq.warranty_expiry_date.strftime("%Y-%m-%d") if eq.warranty_expiry_date else "غير محدد"

            days_rem = "-"
            status_desc = "غير مسجل"
            if eq.warranty_expiry_date:
                delta = (eq.warranty_expiry_date - today).days
                days_rem = str(delta)
                if delta > 30:
                    status_desc = "ساري الضمان"
                elif delta >= 0:
                    status_desc = "يوشك على الانتهاء"
                else:
                    status_desc = f"منتهي ({abs(delta)} يوم)"

            rows.append({
                "id": eq.id,
                "asset_number": eq.asset_number or "-",
                "name_ar": eq.name_ar or "-",
                "model": eq.model or "-",
                "serial_number": eq.serial_number or "-",
                "department": dept_name,
                "vendor": vendor_name,
                "warranty_expiry": w_date_str,
                "days_remaining": days_rem,
                "warranty_status": status_desc
            })

        return {
            "title": "تقرير متابعة عقود الضمان وانتهاء الصلاحية",
            "column_keys": column_keys,
            "summary": {
                "total_count": total_count,
                "under_warranty_count": under_warranty_count,
                "expired_count": expired_count,
                "display": {
                    "إجمالي الأجهزة": total_count,
                    "تحت الضمان الساري": under_warranty_count,
                    "أجهزة منتهية الضمان": expired_count
                }
            },
            "headers": [
                "رقم الأصل",
                "اسم الجهاز",
                "الموديل",
                "الرقم التسلسلي",
                "القسم",
                "الوكيل / المورد",
                "تاريخ نهاية الضمان",
                "الأيام المتبقية",
                "حالة الضمان"
            ],
            "rows": rows
        }

    # ── 10. Periodic Maintenance Summary Report ──────────────────────────────
    def get_periodic_maintenance_report_data(
        self,
        db: Session,
        from_date: Optional[date] = None,
        to_date: Optional[date] = None,
        department_id: Optional[int] = None
    ) -> Dict[str, Any]:
        query = db.query(WorkOrder)
        if from_date:
            query = query.filter(WorkOrder.created_at >= datetime.combine(from_date, datetime.min.time()))
        if to_date:
            query = query.filter(WorkOrder.created_at <= datetime.combine(to_date, datetime.max.time()))
        if department_id:
            query = query.join(Equipment).filter(Equipment.department_id == department_id)

        orders = query.order_by(WorkOrder.created_at.desc()).all()

        total_orders = len(orders)
        completed_orders = sum(1 for o in orders if o.status in ("completed", "closed"))
        total_cost = sum(float(o.total_cost or 0.0) for o in orders)

        column_keys = [
            "work_order_number", "reported_date", "equipment", "department",
            "priority", "status", "technician", "failure_description", "total_cost"
        ]

        rows = []
        for o in orders:
            eq_name = f"{o.equipment.name_ar} ({o.equipment.asset_number})" if o.equipment else "-"
            dept_name = o.equipment.department.name_ar if (o.equipment and o.equipment.department) else "-"
            tech_name = o.assigned_technician.full_name if o.assigned_technician else "-"

            rows.append({
                "id": o.id,
                "work_order_number": o.work_order_number,
                "reported_date": o.created_at.strftime("%Y-%m-%d") if o.created_at else "-",
                "equipment": eq_name,
                "department": dept_name,
                "priority": tr(f"priority.{o.priority}", default=o.priority),
                "status": tr(f"status.{o.status}", default=o.status),
                "technician": tech_name,
                "failure_description": o.failure_description or "-",
                "total_cost": float(o.total_cost or 0.0)
            })

        return {
            "title": "تقرير ملخص أعمال الصيانة الدورية والتكاليف",
            "column_keys": column_keys,
            "summary": {
                "total_orders": total_orders,
                "completed_orders": completed_orders,
                "total_cost": total_cost,
                "display": {
                    "إجمالي البلاغات المسجلة": total_orders,
                    "البلاغات المكتملة": completed_orders,
                    "التكلفة الإجمالية للصيانة": f"{total_cost:,.2f} ر.س"
                }
            },
            "headers": [
                "رقم أمر العمل",
                "تاريخ البلاغ",
                "الجهاز الطبي",
                "القسم",
                "الأولوية",
                "الحالة",
                "الفني المسؤول",
                "وصف العطل",
                "التكلفة (ر.س)"
            ],
            "rows": rows
        }

    # ── 5. Excel Export Function ─────────────────────────────────────────────
    def export_to_excel(self, report_data: Dict[str, Any], filepath: str):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "تقرير MEMMS"
        ws.sheet_view.rightToLeft = True
        try:
            ws.views.sheetView[0].rightToLeft = True
        except Exception:
            pass

        title = report_data.get("title", "تقرير الصيانة والأجهزة الطبية")
        summary = report_data.get("summary", {})
        display_summary = summary.get("display", {k: v for k, v in summary.items() if k != "display"})
        headers = report_data.get("headers", [])
        column_keys = report_data.get("column_keys", [])
        rows = report_data.get("rows", [])

        # Title Block
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(headers), 4))
        title_cell = ws.cell(row=1, column=1, value=f"نظام إدارة صيانة الأجهزة الطبية (MEMMS) — {title}")
        title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Subtitle / Date
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(len(headers), 4))
        sub_cell = ws.cell(row=2, column=1, value=f"تاريخ التقرير: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
        sub_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Summary KPIs Block
        row_idx = 4
        if display_summary:
            ws.cell(row=row_idx, column=1, value="ملخص المؤشرات والإحصائيات:").font = Font(name="Segoe UI", bold=True, size=11, color="0F172A")
            row_idx += 1
            for k, v in display_summary.items():
                val_str = f"{v:,.2f}" if isinstance(v, (int, float)) else str(v)
                c1 = ws.cell(row=row_idx, column=1, value=k)
                c1.font = Font(name="Segoe UI", bold=True, color="475569")
                c2 = ws.cell(row=row_idx, column=2, value=val_str)
                c2.font = Font(name="Segoe UI", bold=True, color="0F766E")
                row_idx += 1
            row_idx += 1

        # Table Headers
        header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx, header_text in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=header_text)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        row_idx += 1

        # Table Rows
        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0")
        )
        alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        for r_i, row in enumerate(rows):
            for col_idx, key in enumerate(column_keys, 1):
                val = row.get(key, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                cell.font = Font(name="Segoe UI", size=10)
                cell.alignment = Alignment(horizontal="center" if isinstance(val, (int, float)) else "right", vertical="center")
                if r_i % 2 == 1:
                    cell.fill = alt_fill
            row_idx += 1

        # Auto-fit column widths
        for col_idx in range(1, len(headers) + 1):
            col_letter = get_column_letter(col_idx)
            header_str = str(headers[col_idx - 1]) if col_idx <= len(headers) else ""
            max_len = len(header_str)
            key = column_keys[col_idx - 1] if col_idx <= len(column_keys) else ""
            for r in rows:
                val_str = str(r.get(key, ""))
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        wb.save(filepath)

    # ── 6. CSV Export Function ───────────────────────────────────────────────
    def export_to_csv(self, report_data: Dict[str, Any], filepath: str):
        headers = report_data.get("headers", [])
        column_keys = report_data.get("column_keys", [])
        rows = report_data.get("rows", [])

        with open(filepath, mode="w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([f"MEMMS — {report_data.get('title', 'التقرير')}"])
            writer.writerow([f"تاريخ التقرير: {datetime.now().strftime('%Y-%m-%d %H:%M')}"])
            writer.writerow([])
            writer.writerow(headers)

            for row in rows:
                row_values = [row.get(k, "") for k in column_keys]
                writer.writerow(row_values)

    def prepare_rtl_document(self, doc, html: str, font_size: int = 9, default_font: str = "Segoe UI"):
        """
        Sets up QTextDocument to strictly format RTL Arabic layouts,
        enforcing Right-to-Left on both paragraph options and all internal QTextTables.
        """
        from PySide6.QtGui import QFont, QTextOption, QTextTable
        from PySide6.QtCore import Qt

        doc.setDefaultFont(QFont(default_font, font_size))
        opt = doc.defaultTextOption()
        opt.setTextDirection(Qt.LayoutDirection.RightToLeft)
        opt.setAlignment(Qt.AlignmentFlag.AlignRight)
        doc.setDefaultTextOption(opt)

        doc.setHtml(html)

        root = doc.rootFrame()
        if root:
            for child in root.childFrames():
                if isinstance(child, QTextTable):
                    fmt = child.format()
                    fmt.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
                    fmt.setAlignment(Qt.AlignmentFlag.AlignRight)
                    child.setFormat(fmt)

    # ── 7. HTML / Printable Document Export ──────────────────────────────────
    def generate_html_report(self, report_data: Dict[str, Any]) -> str:
        """Generates standard HTML template for PDF printing with native Arabic RTL styling."""
        if report_data.get("is_single_equipment"):
            return self.generate_equipment_history_html(report_data)
        if report_data.get("is_work_order_ticket"):
            return self.generate_work_order_ticket_html(report_data)

        title = report_data.get("title", "تقرير الصيانة والأجهزة الطبية")
        summary = report_data.get("summary", {})
        display_summary = summary.get("display", {k: v for k, v in summary.items() if k != "display"})
        headers = report_data.get("headers", [])
        column_keys = report_data.get("column_keys", [])
        rows = report_data.get("rows", [])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        summary_cells = []
        for k, v in display_summary.items():
            val_str = f"{v:,.2f}" if isinstance(v, (int, float)) else str(v)
            summary_cells.append(f"""
            <td dir="rtl" style="background-color: #F8FAFC; border: 1px solid #CBD5E1; padding: 6px 10px; text-align: center;">
                <div style="font-size: 9px; color: #475569; font-weight: bold; margin-bottom: 2px;">{k}</div>
                <div style="font-size: 13px; color: #0F766E; font-weight: bold;">{val_str}</div>
            </td>
            """)

        summary_table_html = ""
        if summary_cells:
            summary_table_html = f"""
            <table dir="rtl" width="100%" style="margin-bottom: 12px; border-collapse: separate; border-spacing: 4px; direction: rtl;">
                <tr>{"".join(summary_cells)}</tr>
            </table>
            """

        headers_html = "".join([
            f"<th style='background-color: #0F766E; color: #FFFFFF; padding: 6pt 4pt; border: 1pt solid #0D9488; font-size: 8.5pt; text-align: center; font-weight: bold;'>{h}</th>"
            for h in headers
        ])

        rows_html = ""
        for idx, r in enumerate(rows):
            bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
            cells = ""
            for k in column_keys:
                val = r.get(k, "-")
                val_str = f"{val:,.2f}" if isinstance(val, float) else str(val)
                align = "center" if isinstance(val, (int, float)) else "right"
                cells += f"<td style='padding: 5pt 5pt; border: 1pt solid #CBD5E1; font-size: 8pt; background-color: {bg}; text-align: {align}; color: #0F172A;'>{val_str}</td>"
            rows_html += f"<tr>{cells}</tr>"

        html = f"""
        <!DOCTYPE html>
        <html dir="rtl" lang="ar">
        <head>
            <meta charset="utf-8">
            <title>{title}</title>
            <style>
                * {{
                    direction: rtl;
                    box-sizing: border-box;
                }}
                body {{
                    font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                    color: #0F172A;
                    direction: rtl;
                    text-align: right;
                    margin: 0;
                    padding: 0;
                }}
                .hospital-title {{
                    font-size: 15pt;
                    font-weight: bold;
                    color: #0F766E;
                    margin: 0 0 2pt 0;
                    text-align: center;
                    direction: rtl;
                }}
                .report-title {{
                    font-size: 12pt;
                    font-weight: bold;
                    color: #1E293B;
                    margin: 0 0 4pt 0;
                    text-align: center;
                    direction: rtl;
                }}
                .header-container {{
                    border-bottom: 2pt solid #0F766E;
                    padding-bottom: 6pt;
                    margin-bottom: 8pt;
                    text-align: center;
                    direction: rtl;
                }}
                .sub {{
                    color: #64748B;
                    font-size: 8pt;
                    font-weight: 600;
                    direction: rtl;
                }}
                table {{
                    direction: rtl;
                    text-align: right;
                }}
                table.data-table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 6pt;
                    direction: rtl;
                }}
                th, td {{
                    direction: rtl;
                }}
                .footer {{
                    margin-top: 10pt;
                    font-size: 7.5pt;
                    color: #94A3B8;
                    text-align: center;
                    border-top: 1pt solid #E2E8F0;
                    padding-top: 4pt;
                    direction: rtl;
                }}
            </style>
        </head>
        <body dir="rtl">
            <div dir="rtl" class="header-container">
                <div class="hospital-title">هيئة مستشفى الثورة العام النموذجي — صنعاء</div>
                <div class="report-title">{title}</div>
                <div class="sub">تاريخ استخراج التقرير: {now_str} &nbsp;|&nbsp; الإدارة العامة للشؤون الهندسية والصيانة الطبية &nbsp;|&nbsp; نظام MEMMS</div>
            </div>
            {summary_table_html}
            <table dir="rtl" class="data-table" width="100%">
                <thead>
                    <tr>{headers_html}</tr>
                </thead>
                <tbody>
                    {rows_html if rows_html else "<tr><td colspan='" + str(len(headers)) + "' style='text-align:center; padding: 14pt; color: #64748B;'>لا توجد بيانات متاحة لهذا التقرير</td></tr>"}
                </tbody>
            </table>
            <div dir="rtl" class="footer">
                تم استخراج هذا التقرير آلياً عبر نظام إدارة صيانة الأجهزة الطبية (MEMMS) — هيئة مستشفى الثورة العام بصنعاء
            </div>
        </body>
        </html>
        """
        return html

    # ── 8. Single Equipment Maintenance History Report ───────────────────────
    def get_equipment_history_report_data(self, db: Session, equipment_id: int) -> Dict[str, Any]:
        """Gathers full lifecycle maintenance history for a single medical asset."""
        eq = db.query(Equipment).filter_by(id=equipment_id).first()
        if not eq:
            return {}

        work_orders = db.query(WorkOrder).filter_by(equipment_id=eq.id).order_by(WorkOrder.created_at.desc()).all()
        pm_schedules = db.query(PMSchedule).filter_by(equipment_id=eq.id).order_by(PMSchedule.due_date.desc()).all()

        wo_ids = [w.id for w in work_orders]
        parts_consumed = []
        if wo_ids:
            parts_consumed = db.query(WorkOrderPart).filter(WorkOrderPart.work_order_id.in_(wo_ids)).all()

        total_wo = len(work_orders)
        completed_wo = sum(1 for w in work_orders if w.status in ("completed", "closed"))
        open_wo = sum(1 for w in work_orders if w.status in ("open", "pending", "assigned", "in_progress", "waiting_parts"))
        total_maint_cost = sum(float(w.total_cost or 0.0) for w in work_orders)
        total_parts_cost = sum(float(w.total_parts_cost or 0.0) for w in work_orders)
        total_labor_cost = sum(float(w.total_labor_cost or 0.0) for w in work_orders)

        cost_ratio = (total_maint_cost / float(eq.purchase_price) * 100) if (eq.purchase_price and eq.purchase_price > 0) else 0.0

        wo_rows = []
        for w in work_orders:
            wo_rows.append({
                "work_order_number": w.work_order_number,
                "created_at": w.created_at.strftime("%Y-%m-%d") if w.created_at else "-",
                "priority": tr(f"priority.{w.priority}", default=w.priority),
                "status": tr(f"status.{w.status}", default=w.status),
                "technician": w.assigned_technician.full_name if w.assigned_technician else "غير مخصص",
                "failure_description": w.failure_description or "-",
                "resolution_notes": w.corrective_action or w.diagnosis or "-",
                "total_cost": float(w.total_cost or 0.0)
            })

        pm_rows = []
        for pm in pm_schedules:
            tech_name = pm.execution.technician.full_name if (pm.execution and pm.execution.technician) else "غير مخصص"
            plan_title = pm.plan.title if (pm.plan and hasattr(pm.plan, "title")) else (getattr(pm.plan, "name_ar", "-") if pm.plan else "-")
            pm_rows.append({
                "schedule_code": f"PM-{pm.id:04d}",
                "plan_name": plan_title,
                "scheduled_date": pm.due_date.strftime("%Y-%m-%d") if pm.due_date else "-",
                "status": tr(f"status.{pm.status}", default=pm.status),
                "technician": tech_name,
            })

        part_rows = []
        for wp in parts_consumed:
            p_name = wp.part.name_ar if (wp.part and wp.part.name_ar) else (wp.part.name_en if wp.part else "-")
            part_rows.append({
                "part_number": wp.part.part_number if wp.part else "-",
                "name": p_name,
                "quantity": float(wp.quantity or 0.0),
                "unit_price": float(wp.unit_price or 0.0),
                "total_price": float(wp.total_price or 0.0),
            })

        dept_name = eq.department.name_ar if eq.department else "-"
        loc_name = eq.location.room_name if eq.location else "-"
        manuf_name = eq.manufacturer.name if eq.manufacturer else "-"
        agent_name = eq.agent.name if eq.agent else "-"

        criticality_text = tr(f"criticality.{eq.criticality}", default=eq.criticality)
        risk_text = tr(f"risk.{eq.risk_level}", default=eq.risk_level)
        status_text = tr(f"status.{eq.status}", default=eq.status)

        column_keys = ["work_order_number", "created_at", "priority", "status", "technician", "failure_description", "total_cost"]
        headers = ["رقم أمر العمل", "التاريخ", "الأولوية", "الحالة", "الفني المكلف", "وصف العطل", "التكلفة (ر.س)"]

        return {
            "title": f"بطاقة السجل الفني وتاريخ صيانة الجهاز: {eq.name_ar or eq.name_en} ({eq.asset_number})",
            "is_single_equipment": True,
            "equipment": {
                "id": eq.id,
                "asset_number": eq.asset_number,
                "serial_number": eq.serial_number,
                "name_ar": eq.name_ar or "-",
                "name_en": eq.name_en or "-",
                "model": eq.model or "-",
                "category": eq.category or "-",
                "department": dept_name,
                "location": loc_name,
                "manufacturer": manuf_name,
                "agent": agent_name,
                "status": status_text,
                "risk_level": risk_text,
                "criticality": criticality_text,
                "purchase_date": eq.purchase_date.strftime("%Y-%m-%d") if eq.purchase_date else "-",
                "installation_date": eq.installation_date.strftime("%Y-%m-%d") if eq.installation_date else "-",
                "warranty_expiry": eq.warranty_expiry_date.strftime("%Y-%m-%d") if eq.warranty_expiry_date else "-",
                "purchase_price": float(eq.purchase_price or 0.0),
                "notes": eq.notes or "-",
            },
            "summary": {
                "total_orders": total_wo,
                "completed_orders": completed_wo,
                "open_orders": open_wo,
                "total_maint_cost": total_maint_cost,
                "total_parts_cost": total_parts_cost,
                "total_labor_cost": total_labor_cost,
                "cost_ratio": cost_ratio,
                "display": {
                    "إجمالي أوامر الصيانة": total_wo,
                    "البلاغات المكتملة": completed_wo,
                    "البلاغات المفتوحة": open_wo,
                    "إجمالي تكاليف الصيانة": f"{total_maint_cost:,.2f} ر.س",
                    "نسبة الصيانة لسعر الشراء": f"{cost_ratio:.1f}%"
                }
            },
            "headers": headers,
            "column_keys": column_keys,
            "rows": wo_rows,
            "work_orders": wo_rows,
            "pm_schedules": pm_rows,
            "parts_consumed": part_rows
        }

    def generate_equipment_history_html(self, report_data: Dict[str, Any]) -> str:
        """Generates dedicated Medical Equipment Passport & Maintenance History HTML report."""
        eq = report_data.get("equipment", {})
        summary = report_data.get("summary", {})
        wo_list = report_data.get("work_orders", [])
        pm_list = report_data.get("pm_schedules", [])
        parts_list = report_data.get("parts_consumed", [])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Work Orders Table rows
        wo_rows_html = ""
        for idx, w in enumerate(wo_list):
            bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
            wo_rows_html += f"""
            <tr style="background-color: {bg};">
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center; font-weight: bold; color: #0F766E;">{w.get('work_order_number')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{w.get('created_at')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{w.get('priority')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{w.get('status')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: right;">{w.get('technician')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: right;">{w.get('failure_description')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center; font-weight: bold;">{w.get('total_cost', 0):,.2f} ر.س</td>
            </tr>
            """
        if not wo_rows_html:
            wo_rows_html = "<tr><td colspan='7' style='text-align: center; padding: 10pt; color: #64748B;'>لا توجد بلاغات أو أوامر صيانة مسجلة لهذا الجهاز حتى الآن</td></tr>"

        # PM Schedules rows
        pm_rows_html = ""
        for idx, pm in enumerate(pm_list):
            bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
            pm_rows_html += f"""
            <tr style="background-color: {bg};">
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center; font-weight: bold;">{pm.get('schedule_code')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: right;">{pm.get('plan_name')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{pm.get('scheduled_date')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{pm.get('status')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: right;">{pm.get('technician')}</td>
            </tr>
            """
        if not pm_rows_html:
            pm_rows_html = "<tr><td colspan='5' style='text-align: center; padding: 8pt; color: #64748B;'>لا توجد جداول صيانة وقائية مرتبطة بهذا الجهاز</td></tr>"

        # Parts Consumed rows
        parts_rows_html = ""
        for idx, p in enumerate(parts_list):
            bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
            parts_rows_html += f"""
            <tr style="background-color: {bg};">
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{p.get('part_number')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: right;">{p.get('name')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{p.get('quantity'):g}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{p.get('unit_price', 0):,.2f} ر.س</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center; font-weight: bold;">{p.get('total_price', 0):,.2f} ر.س</td>
            </tr>
            """
        if not parts_rows_html:
            parts_rows_html = "<tr><td colspan='5' style='text-align: center; padding: 8pt; color: #64748B;'>لم يتم تسجيل صرف أي قطع غيار على هذا الجهاز</td></tr>"

        html = f"""
        <!DOCTYPE html>
        <html dir="rtl" lang="ar">
        <head>
            <meta charset="utf-8">
            <title>بطاقة السجل الفني وتاريخ صيانة الجهاز</title>
            <style>
                * {{
                    direction: rtl;
                    box-sizing: border-box;
                }}
                body {{
                    font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                    color: #0F172A;
                    direction: rtl;
                    text-align: right;
                    margin: 0;
                    padding: 0;
                }}
                .header-box {{
                    border-bottom: 2.5pt solid #0F766E;
                    padding-bottom: 8pt;
                    margin-bottom: 10pt;
                    text-align: center;
                    direction: rtl;
                }}
                .hospital-name {{
                    font-size: 15pt;
                    font-weight: 800;
                    color: #0F766E;
                    margin: 0 0 2pt 0;
                    direction: rtl;
                }}
                .report-title {{
                    font-size: 13pt;
                    font-weight: 700;
                    color: #1E293B;
                    margin: 2pt 0 4pt 0;
                    direction: rtl;
                }}
                .sub {{
                    color: #64748B;
                    font-size: 8.5pt;
                    font-weight: 600;
                    direction: rtl;
                }}
                .section-header {{
                    background-color: #0F766E;
                    color: #FFFFFF;
                    font-size: 9.5pt;
                    font-weight: bold;
                    padding: 4pt 8pt;
                    border-radius: 3pt;
                    margin-top: 10pt;
                    margin-bottom: 5pt;
                    direction: rtl;
                    text-align: right;
                }}
                table {{
                    direction: rtl;
                    text-align: right;
                }}
                table.info-table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 6pt;
                    direction: rtl;
                }}
                table.info-table td {{
                    padding: 4.5pt 6pt;
                    border: 1pt solid #E2E8F0;
                    font-size: 8.5pt;
                    direction: rtl;
                }}
                table.info-table td.lbl {{
                    background-color: #F8FAFC;
                    color: #475569;
                    font-weight: 700;
                    width: 18%;
                    text-align: right;
                }}
                table.info-table td.val {{
                    color: #0F172A;
                    width: 32%;
                    text-align: right;
                }}
                .kpi-table {{
                    width: 100%;
                    border-collapse: separate;
                    border-spacing: 4pt;
                    margin: 6pt 0;
                    direction: rtl;
                }}
                .kpi-cell {{
                    background-color: #F8FAFC;
                    border: 1pt solid #CBD5E1;
                    padding: 5pt 8pt;
                    text-align: center;
                    border-radius: 4pt;
                    direction: rtl;
                }}
                table.data-table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 6pt;
                    direction: rtl;
                }}
                table.data-table th {{
                    background-color: #0D9488;
                    color: #FFFFFF;
                    padding: 5pt 4pt;
                    border: 1pt solid #0F766E;
                    font-size: 8.5pt;
                    text-align: center;
                    direction: rtl;
                }}
                th, td {{
                    direction: rtl;
                }}
                .signatures {{
                    margin-top: 16pt;
                    width: 100%;
                    border-collapse: collapse;
                    direction: rtl;
                }}
                .sig-box {{
                    width: 33%;
                    text-align: center;
                    padding: 8pt;
                    border-top: 1pt solid #94A3B8;
                    font-size: 8.5pt;
                    color: #334155;
                    direction: rtl;
                }}
                .footer {{
                    margin-top: 14pt;
                    font-size: 7.5pt;
                    color: #94A3B8;
                    text-align: center;
                    border-top: 1pt solid #E2E8F0;
                    padding-top: 4pt;
                    direction: rtl;
                }}
            </style>
        </head>
        <body dir="rtl">
            <div dir="rtl" class="header-box">
                <div class="hospital-name">هيئة مستشفى الثورة العام النموذجي — صنعاء</div>
                <div class="report-title">بطاقة السجل الفني وتاريخ صيانة الجهاز الطبي (Equipment Passport)</div>
                <div class="sub">الإدارة العامة للشؤون الهندسية والصيانة الطبية &nbsp;|&nbsp; استخراج النظام: {now_str}</div>
            </div>

            <div dir="rtl" class="section-header">1. البيانات التعريفية والمواصفات السريرية للجهاز</div>
            <table dir="rtl" class="info-table">
                <tr>
                    <td class="lbl">رقم الأصل (Asset ID):</td>
                    <td class="val" style="font-weight: bold; color: #0F766E;">{eq.get('asset_number')}</td>
                    <td class="lbl">الرقم التسلسلي (S/N):</td>
                    <td class="val" style="font-weight: bold;">{eq.get('serial_number')}</td>
                </tr>
                <tr>
                    <td class="lbl">اسم الجهاز الطبي:</td>
                    <td class="val" style="font-weight: bold;">{eq.get('name_ar')}</td>
                    <td class="lbl">الاسم بالإنجليزية والموديل:</td>
                    <td class="val">{eq.get('name_en')} ({eq.get('model')})</td>
                </tr>
                <tr>
                    <td class="lbl">القسم الطبي والموقع:</td>
                    <td class="val">{eq.get('department')} — {eq.get('location')}</td>
                    <td class="lbl">التصنيف الطبي:</td>
                    <td class="val">{eq.get('category')}</td>
                </tr>
                <tr>
                    <td class="lbl">الشركة المصنعة:</td>
                    <td class="val">{eq.get('manufacturer')}</td>
                    <td class="lbl">الوكيل المحلي المعتمد:</td>
                    <td class="val">{eq.get('agent')}</td>
                </tr>
                <tr>
                    <td class="lbl">الحالة التشغيلية الحالية:</td>
                    <td class="val" style="font-weight: bold;">{eq.get('status')}</td>
                    <td class="lbl">درجة الخطورة:</td>
                    <td class="val">{eq.get('risk_level')}</td>
                </tr>
                <tr>
                    <td class="lbl">درجة الحرجية السريرية:</td>
                    <td class="val" style="font-weight: bold; color: #0F766E;">{eq.get('criticality')}</td>
                    <td class="lbl">تاريخ الشراء والتركيب:</td>
                    <td class="val">{eq.get('purchase_date')} / {eq.get('installation_date')}</td>
                </tr>
                <tr>
                    <td class="lbl">سعر الشراء الأساسي:</td>
                    <td class="val" style="font-weight: bold;">{eq.get('purchase_price', 0):,.2f} ر.س</td>
                    <td class="lbl">تاريخ انتهاء فترة الضمان:</td>
                    <td class="val">{eq.get('warranty_expiry')}</td>
                </tr>
            </table>

            <div dir="rtl" class="section-header">2. مؤشرات الأداء والتكاليف التراكمية لصيانة الجهاز</div>
            <table dir="rtl" class="kpi-table">
                <tr>
                    <td class="kpi-cell">
                        <div style="font-size: 8pt; color: #64748B;">إجمالي أوامر الصيانة</div>
                        <div style="font-size: 13pt; font-weight: bold; color: #0F766E;">{summary.get('total_orders', 0)}</div>
                    </td>
                    <td class="kpi-cell">
                        <div style="font-size: 8pt; color: #64748B;">البلاغات المكتملة</div>
                        <div style="font-size: 13pt; font-weight: bold; color: #059669;">{summary.get('completed_orders', 0)}</div>
                    </td>
                    <td class="kpi-cell">
                        <div style="font-size: 8pt; color: #64748B;">البلاغات الجارية</div>
                        <div style="font-size: 13pt; font-weight: bold; color: #D97706;">{summary.get('open_orders', 0)}</div>
                    </td>
                    <td class="kpi-cell">
                        <div style="font-size: 8pt; color: #64748B;">تكلفة قطع الغيار المستهلكة</div>
                        <div style="font-size: 12pt; font-weight: bold; color: #475569;">{summary.get('total_parts_cost', 0):,.2f} ر.س</div>
                    </td>
                    <td class="kpi-cell">
                        <div style="font-size: 8pt; color: #64748B;">إجمالي تكاليف الصيانة</div>
                        <div style="font-size: 12pt; font-weight: bold; color: #DC2626;">{summary.get('total_maint_cost', 0):,.2f} ر.س</div>
                    </td>
                    <td class="kpi-cell">
                        <div style="font-size: 8pt; color: #64748B;">نسبة الصيانة لسعر الشراء</div>
                        <div style="font-size: 13pt; font-weight: bold; color: #0F766E;">{summary.get('cost_ratio', 0):.1f}%</div>
                    </td>
                </tr>
            </table>

            <div dir="rtl" class="section-header">3. السجل التاريخي لأوامر وبلاغات الصيانة المنفذة (Work Orders)</div>
            <table dir="rtl" class="data-table">
                <thead>
                    <tr>
                        <th style="width: 13%;">رقم أمر العمل</th>
                        <th style="width: 10%;">التاريخ</th>
                        <th style="width: 9%;">الأولوية</th>
                        <th style="width: 10%;">الحالة</th>
                        <th style="width: 15%;">الفني المكلف</th>
                        <th>وصف العطل / المشكلة</th>
                        <th style="width: 12%;">التكلفة</th>
                    </tr>
                </thead>
                <tbody>
                    {wo_rows_html}
                </tbody>
            </table>

            <div dir="rtl" class="section-header">4. سجل الصيانة الوقائية والفحص الدوري المجدول (PM Plans)</div>
            <table dir="rtl" class="data-table">
                <thead>
                    <tr>
                        <th style="width: 15%;">كود الجدول</th>
                        <th>خطة الصيانة الوقائية</th>
                        <th style="width: 15%;">التاريخ المجدول</th>
                        <th style="width: 15%;">الحالة</th>
                        <th style="width: 20%;">الفني المكلف</th>
                    </tr>
                </thead>
                <tbody>
                    {pm_rows_html}
                </tbody>
            </table>

            <div dir="rtl" class="section-header">5. قطع الغيار المستبدلة والمستهلكة على الجهاز</div>
            <table dir="rtl" class="data-table">
                <thead>
                    <tr>
                        <th style="width: 18%;">رقم القطعة / الكود</th>
                        <th>اسم قطعة الغيار</th>
                        <th style="width: 12%;">الكمية</th>
                        <th style="width: 16%;">سعر الوحدة</th>
                        <th style="width: 18%;">الإجمالي</th>
                    </tr>
                </thead>
                <tbody>
                    {parts_rows_html}
                </tbody>
            </table>

            <table dir="rtl" class="signatures">
                <tr>
                    <td class="sig-box">
                        <strong>مهندس الصيانة الطبية المسؤول</strong><br><br>
                        التوقيع: ............................
                    </td>
                    <td class="sig-box">
                        <strong>رئيس قسم الهندسة الطبية</strong><br><br>
                        التوقيع: ............................
                    </td>
                    <td class="sig-box">
                        <strong>مدير عام المستشفى / الاعتماد</strong><br><br>
                        الختم الرسمي: ............................
                    </td>
                </tr>
            </table>

            <div dir="rtl" class="footer">
                نظام إدارة صيانة الأجهزة الطبية (MEMMS) — هيئة مستشفى الثورة العام النموذجي بصنعاء — السجل التاريخي للأصل الطبي
            </div>
        </body>
        </html>
        """
        return html

    # ── 9. Individual Work Order Maintenance Ticket / Job Sheet ──────────────
    def get_work_order_ticket_data(self, db: Session, work_order_id: int) -> Dict[str, Any]:
        """Gathers single work order ticket details with parts and equipment profile."""
        wo = db.query(WorkOrder).filter_by(id=work_order_id).first()
        if not wo:
            return {}

        eq = wo.equipment
        parts = db.query(WorkOrderPart).filter_by(work_order_id=wo.id).all()

        part_rows = []
        for p in parts:
            p_name = p.part.name_ar if (p.part and p.part.name_ar) else (p.part.name_en if p.part else "-")
            part_rows.append({
                "part_number": p.part.part_number if p.part else "-",
                "name": p_name,
                "quantity": float(p.quantity or 0.0),
                "unit_price": float(p.unit_price or 0.0),
                "total_price": float(p.total_price or 0.0)
            })

        return {
            "title": f"إشعار وتذكرة أمر صيانة رقم {wo.work_order_number}",
            "is_work_order_ticket": True,
            "work_order": {
                "id": wo.id,
                "number": wo.work_order_number,
                "created_at": wo.created_at.strftime("%Y-%m-%d %H:%M") if wo.created_at else "-",
                "completed_at": wo.completed_at.strftime("%Y-%m-%d %H:%M") if wo.completed_at else "-",
                "priority": tr(f"priority.{wo.priority}", default=wo.priority),
                "status": tr(f"status.{wo.status}", default=wo.status),
                "reported_by": wo.reported_by_name or "-",
                "department": wo.department.name_ar if wo.department else (eq.department.name_ar if eq and eq.department else "-"),
                "technician": wo.assigned_technician.full_name if wo.assigned_technician else "غير مخصص",
                "failure_description": wo.failure_description or "-",
                "resolution_notes": wo.corrective_action or wo.diagnosis or "-",
                "labor_cost": float(wo.total_labor_cost or 0.0),
                "parts_cost": float(wo.total_parts_cost or 0.0),
                "total_cost": float(wo.total_cost or 0.0)
            },
            "equipment": {
                "asset_number": eq.asset_number if eq else "-",
                "serial_number": eq.serial_number if eq else "-",
                "name_ar": eq.name_ar if eq else "-",
                "model": eq.model if eq else "-",
                "criticality": tr(f"criticality.{eq.criticality}", default=eq.criticality) if eq else "-",
                "risk_level": tr(f"risk.{eq.risk_level}", default=eq.risk_level) if eq else "-"
            } if eq else {},
            "parts": part_rows
        }

    def generate_work_order_ticket_html(self, report_data: Dict[str, Any]) -> str:
        """Generates printable Work Order Maintenance Ticket / Handover Sheet."""
        wo = report_data.get("work_order", {})
        eq = report_data.get("equipment", {})
        parts = report_data.get("parts", [])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        parts_rows_html = ""
        for idx, p in enumerate(parts):
            bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
            parts_rows_html += f"""
            <tr style="background-color: {bg};">
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{p.get('part_number')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: right;">{p.get('name')}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{p.get('quantity'):g}</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center;">{p.get('unit_price', 0):,.2f} ر.س</td>
                <td style="padding: 5pt; border: 1pt solid #CBD5E1; text-align: center; font-weight: bold;">{p.get('total_price', 0):,.2f} ر.س</td>
            </tr>
            """
        if not parts_rows_html:
            parts_rows_html = "<tr><td colspan='5' style='text-align: center; padding: 6pt; color: #64748B;'>لم يتم استهلاك قطع غيار في أمر الصيانة هذا</td></tr>"

        html = f"""
        <!DOCTYPE html>
        <html dir="rtl" lang="ar">
        <head>
            <meta charset="utf-8">
            <title>تذكرة أمر صيانة رقم {wo.get('number')}</title>
            <style>
                * {{
                    direction: rtl;
                    box-sizing: border-box;
                }}
                body {{
                    font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                    color: #0F172A;
                    direction: rtl;
                    text-align: right;
                    margin: 0;
                    padding: 0;
                }}
                .header-box {{
                    border-bottom: 2.5pt solid #0284C7;
                    padding-bottom: 8pt;
                    margin-bottom: 10pt;
                    text-align: center;
                    direction: rtl;
                }}
                .hospital-name {{
                    font-size: 15pt;
                    font-weight: 800;
                    color: #0284C7;
                    margin: 0 0 2pt 0;
                    direction: rtl;
                }}
                .report-title {{
                    font-size: 13pt;
                    font-weight: 700;
                    color: #1E293B;
                    margin: 2pt 0 4pt 0;
                    direction: rtl;
                }}
                .sub {{
                    color: #64748B;
                    font-size: 8.5pt;
                    font-weight: 600;
                    direction: rtl;
                }}
                .section-header {{
                    background-color: #0284C7;
                    color: #FFFFFF;
                    font-size: 9.5pt;
                    font-weight: bold;
                    padding: 4pt 8pt;
                    border-radius: 3pt;
                    margin-top: 10pt;
                    margin-bottom: 5pt;
                    direction: rtl;
                    text-align: right;
                }}
                table {{
                    direction: rtl;
                    text-align: right;
                }}
                table.info-table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 6pt;
                    direction: rtl;
                }}
                table.info-table td {{
                    padding: 5pt 7pt;
                    border: 1pt solid #E2E8F0;
                    font-size: 8.5pt;
                    direction: rtl;
                }}
                table.info-table td.lbl {{
                    background-color: #F8FAFC;
                    color: #475569;
                    font-weight: 700;
                    width: 20%;
                    text-align: right;
                }}
                table.info-table td.val {{
                    color: #0F172A;
                    width: 30%;
                    text-align: right;
                }}
                table.data-table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 6pt;
                    direction: rtl;
                }}
                table.data-table th {{
                    background-color: #0369A1;
                    color: #FFFFFF;
                    padding: 5pt 4pt;
                    border: 1pt solid #0284C7;
                    font-size: 8.5pt;
                    text-align: center;
                    direction: rtl;
                }}
                th, td {{
                    direction: rtl;
                }}
                .desc-box {{
                    border: 1pt solid #CBD5E1;
                    background-color: #F8FAFC;
                    padding: 8pt;
                    border-radius: 4pt;
                    font-size: 8.5pt;
                    line-height: 1.5;
                    margin-bottom: 6pt;
                    direction: rtl;
                    text-align: right;
                }}
                .signatures {{
                    margin-top: 18pt;
                    width: 100%;
                    border-collapse: collapse;
                    direction: rtl;
                }}
                .sig-box {{
                    width: 33%;
                    text-align: center;
                    padding: 8pt;
                    border-top: 1pt solid #94A3B8;
                    font-size: 8.5pt;
                    color: #334155;
                    direction: rtl;
                }}
                .footer {{
                    margin-top: 14pt;
                    font-size: 7.5pt;
                    color: #94A3B8;
                    text-align: center;
                    border-top: 1pt solid #E2E8F0;
                    padding-top: 4pt;
                    direction: rtl;
                }}
            </style>
        </head>
        <body dir="rtl">
            <div dir="rtl" class="header-box">
                <div class="hospital-name">هيئة مستشفى الثورة العام النموذجي — صنعاء</div>
                <div class="report-title">إشعار وتذكرة أمر صيانة طبية (Work Order Maintenance Ticket)</div>
                <div class="sub">الإدارة العامة للشؤون الهندسية والصيانة الطبية &nbsp;|&nbsp; تاريخ الطباعة: {now_str}</div>
            </div>

            <div dir="rtl" class="section-header">1. بيانات أمر الصيانة والبلاغ</div>
            <table dir="rtl" class="info-table">
                <tr>
                    <td class="lbl">رقم أمر العمل (WO #):</td>
                    <td class="val" style="font-weight: bold; color: #0284C7; font-size: 10pt;">{wo.get('number')}</td>
                    <td class="lbl">تاريخ ووقت البلاغ:</td>
                    <td class="val">{wo.get('created_at')}</td>
                </tr>
                <tr>
                    <td class="lbl">درجة الأولوية:</td>
                    <td class="val" style="font-weight: bold;">{wo.get('priority')}</td>
                    <td class="lbl">حالة أمر العمل:</td>
                    <td class="val" style="font-weight: bold;">{wo.get('status')}</td>
                </tr>
                <tr>
                    <td class="lbl">القسم الطبي الطالب:</td>
                    <td class="val">{wo.get('department')}</td>
                    <td class="lbl">اسم مقدم البلاغ:</td>
                    <td class="val">{wo.get('reported_by')}</td>
                </tr>
                <tr>
                    <td class="lbl">الفني / المهندس المكلف:</td>
                    <td class="val" style="font-weight: bold;">{wo.get('technician')}</td>
                    <td class="lbl">تاريخ إتمام الصيانة:</td>
                    <td class="val">{wo.get('completed_at')}</td>
                </tr>
            </table>

            <div dir="rtl" class="section-header">2. بيانات الجهاز الطبي الخاضع للصيانة</div>
            <table dir="rtl" class="info-table">
                <tr>
                    <td class="lbl">رقم الأصل (Asset Tag):</td>
                    <td class="val" style="font-weight: bold; color: #0F766E;">{eq.get('asset_number')}</td>
                    <td class="lbl">الرقم التسلسلي (S/N):</td>
                    <td class="val">{eq.get('serial_number')}</td>
                </tr>
                <tr>
                    <td class="lbl">اسم الجهاز والموديل:</td>
                    <td class="val" style="font-weight: bold;">{eq.get('name_ar')} ({eq.get('model')})</td>
                    <td class="lbl">درجة الحرجية السريرية:</td>
                    <td class="val">{eq.get('criticality')}</td>
                </tr>
            </table>

            <div dir="rtl" class="section-header">3. وصف العطل وبلاغ المشكلة</div>
            <div dir="rtl" class="desc-box">
                {wo.get('failure_description')}
            </div>

            <div dir="rtl" class="section-header">4. الإجراءات الهندسية والتصحيحية المنفذة</div>
            <div dir="rtl" class="desc-box">
                {wo.get('resolution_notes')}
            </div>

            <div dir="rtl" class="section-header">5. قطع الغيار المستهلكة وتكاليف الصيانة</div>
            <table dir="rtl" class="data-table">
                <thead>
                    <tr>
                        <th style="width: 20%;">رقم القطعة</th>
                        <th>اسم قطعة الغيار</th>
                        <th style="width: 12%;">الكمية</th>
                        <th style="width: 16%;">سعر الوحدة</th>
                        <th style="width: 18%;">الإجمالي</th>
                    </tr>
                </thead>
                <tbody>
                    {parts_rows_html}
                </tbody>
            </table>

            <table dir="rtl" class="info-table" style="margin-top: 6pt;">
                <tr>
                    <td class="lbl" style="text-align: center;">تكلفة العمالة المباشرة:</td>
                    <td class="val" style="text-align: center; font-weight: bold;">{wo.get('labor_cost', 0):,.2f} ر.س</td>
                    <td class="lbl" style="text-align: center;">تكلفة قطع الغيار:</td>
                    <td class="val" style="text-align: center; font-weight: bold;">{wo.get('parts_cost', 0):,.2f} ر.س</td>
                    <td class="lbl" style="text-align: center; background-color: #E0F2FE; color: #0369A1;">إجمالي تكلفة أمر الصيانة:</td>
                    <td class="val" style="text-align: center; font-weight: bold; color: #0284C7; font-size: 10pt;">{wo.get('total_cost', 0):,.2f} ر.س</td>
                </tr>
            </table>

            <table dir="rtl" class="signatures">
                <tr>
                    <td class="sig-box">
                        <strong>مهندس / فني الصيانة المنفذ</strong><br><br>
                        الاسم: {wo.get('technician')}<br>
                        التوقيع: ............................
                    </td>
                    <td class="sig-box">
                        <strong>مستلم الجهاز بالقسم الطبي</strong><br><br>
                        الاسم: ............................<br>
                        التوقيع: ............................
                    </td>
                    <td class="sig-box">
                        <strong>اعتماد رئيس قسم الهندسة الطبية</strong><br><br>
                        التوقيع: ............................<br>
                        الختم: ............................
                    </td>
                </tr>
            </table>

            <div dir="rtl" class="footer">
                نظام إدارة صيانة الأجهزة الطبية (MEMMS) — هيئة مستشفى الثورة العام بصنعاء — إشعار صيانة رسمي
            </div>
        </body>
        </html>
        """
        return html


report_service = ReportService()

