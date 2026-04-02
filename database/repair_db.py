import sqlite3
import datetime
from typing import List, Optional
from models.repair_models import (
    Supplier, Agreement, Material, Purchase, PurchaseItem, MaterialBatch,
    Customer, Order, WorkItem, OrderMaterial, OrderMaterialBatch, Document
)
from utils.file_utils import get_db_path

DB_PATH = get_db_path('repair.db')

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # Таблица поставщиков
    c.execute('''CREATE TABLE IF NOT EXISTS suppliers
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT NOT NULL,
                  phone TEXT,
                  email TEXT,
                  address TEXT,
                  notes TEXT)''')
    # Таблица договоров
    c.execute('''CREATE TABLE IF NOT EXISTS agreements
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  supplier_id INTEGER NOT NULL,
                  agreement_number TEXT NOT NULL,
                  date TEXT NOT NULL,
                  total_amount REAL,
                  spent_amount REAL DEFAULT 0,
                  status TEXT,
                  notes TEXT,
                  FOREIGN KEY(supplier_id) REFERENCES suppliers(id))''')
    # Таблица материалов
    c.execute('''CREATE TABLE IF NOT EXISTS materials
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT NOT NULL,
                  unit TEXT,
                  purchase_price REAL,
                  sale_price REAL,
                  stock REAL,
                  inventory_number TEXT,
                  type TEXT DEFAULT 'материал')''')
    # Проверка наличия колонки type
    try:
        c.execute("SELECT type FROM materials LIMIT 1")
    except sqlite3.OperationalError:
        c.execute("ALTER TABLE materials ADD COLUMN type TEXT DEFAULT 'материал'")
    # Таблица счетов
    c.execute('''CREATE TABLE IF NOT EXISTS purchases
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  agreement_id INTEGER NOT NULL,
                  date TEXT NOT NULL,
                  invoice_number TEXT,
                  notes TEXT,
                  FOREIGN KEY(agreement_id) REFERENCES agreements(id))''')
    # Таблица позиций счетов
    c.execute('''CREATE TABLE IF NOT EXISTS purchase_items
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  purchase_id INTEGER NOT NULL,
                  material_id INTEGER NOT NULL,
                  quantity REAL,
                  purchase_price REAL,
                  FOREIGN KEY(purchase_id) REFERENCES purchases(id),
                  FOREIGN KEY(material_id) REFERENCES materials(id))''')
    # Таблица партий материалов
    c.execute('''CREATE TABLE IF NOT EXISTS material_batches
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  purchase_id INTEGER NOT NULL,
                  material_id INTEGER NOT NULL,
                  quantity REAL,
                  purchase_price REAL,
                  date TEXT NOT NULL,
                  batch_number TEXT,
                  FOREIGN KEY(purchase_id) REFERENCES purchases(id),
                  FOREIGN KEY(material_id) REFERENCES materials(id))''')
    # Таблица заказчиков
    c.execute('''CREATE TABLE IF NOT EXISTS customers
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT NOT NULL,
                  phone TEXT,
                  email TEXT,
                  address TEXT,
                  notes TEXT)''')
    # Таблица заказов
    c.execute('''CREATE TABLE IF NOT EXISTS orders
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  customer_id INTEGER NOT NULL,
                  customer_name TEXT,
                  order_number TEXT UNIQUE NOT NULL,
                  date TEXT NOT NULL,
                  status TEXT,
                  notes TEXT,
                  FOREIGN KEY(customer_id) REFERENCES customers(id))''')
    # Таблица работ
    c.execute('''CREATE TABLE IF NOT EXISTS work_items
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  order_id INTEGER NOT NULL,
                  description TEXT,
                  hours REAL,
                  price REAL,
                  FOREIGN KEY(order_id) REFERENCES orders(id))''')
    # Таблица использованных материалов
    c.execute('''CREATE TABLE IF NOT EXISTS order_materials
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  work_item_id INTEGER NOT NULL,
                  material_id INTEGER NOT NULL,
                  quantity REAL,
                  sale_price REAL,
                  FOREIGN KEY(work_item_id) REFERENCES work_items(id),
                  FOREIGN KEY(material_id) REFERENCES materials(id))''')
    # Таблица связей с партиями
    c.execute('''CREATE TABLE IF NOT EXISTS order_material_batches
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  order_material_id INTEGER NOT NULL,
                  material_batch_id INTEGER NOT NULL,
                  quantity REAL,
                  FOREIGN KEY(order_material_id) REFERENCES order_materials(id),
                  FOREIGN KEY(material_batch_id) REFERENCES material_batches(id))''')
    # Таблица документов
    c.execute('''CREATE TABLE IF NOT EXISTS documents
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  supplier_id INTEGER,
                  agreement_id INTEGER,
                  purchase_id INTEGER,
                  document_type TEXT,
                  file_path TEXT NOT NULL,
                  file_name TEXT,
                  uploaded_date TEXT,
                  description TEXT)''')
    conn.commit()
    conn.close()

# ---------- Поставщики ----------
def get_all_suppliers() -> List[Supplier]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, name, phone, email, address, notes FROM suppliers ORDER BY name")
    rows = c.fetchall()
    conn.close()
    return [Supplier(*row) for row in rows]

def get_supplier_by_id(supplier_id: int) -> Optional[Supplier]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, name, phone, email, address, notes FROM suppliers WHERE id=?", (supplier_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return Supplier(*row)
    return None

def add_supplier(supplier: Supplier) -> int:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT INTO suppliers (name, phone, email, address, notes) VALUES (?,?,?,?,?)",
              (supplier.name, supplier.phone, supplier.email, supplier.address, supplier.notes))
    conn.commit()
    sid = c.lastrowid
    conn.close()
    return sid

def update_supplier(supplier: Supplier):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE suppliers SET name=?, phone=?, email=?, address=?, notes=? WHERE id=?",
              (supplier.name, supplier.phone, supplier.email, supplier.address, supplier.notes, supplier.id))
    conn.commit()
    conn.close()

def delete_supplier(supplier_id: int):
    if has_supplier_materials(supplier_id):
        raise ValueError("Нельзя удалить поставщика, у которого есть материалы")
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM suppliers WHERE id=?", (supplier_id,))
    conn.commit()
    conn.close()

# ---------- Договоры ----------
def get_agreements_by_supplier(supplier_id: int) -> List[Agreement]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        SELECT a.id, a.supplier_id, s.name, a.agreement_number, a.date,
               a.total_amount, a.spent_amount, a.status, a.notes
        FROM agreements a
        LEFT JOIN suppliers s ON a.supplier_id = s.id
        WHERE a.supplier_id = ?
        ORDER BY a.date DESC
    """, (supplier_id,))
    rows = c.fetchall()
    conn.close()
    result = []
    for row in rows:
        date_obj = datetime.date.fromisoformat(row[4])
        result.append(Agreement(
            id=row[0], supplier_id=row[1], supplier_name=row[2] or "",
            agreement_number=row[3], date=date_obj, total_amount=row[5],
            spent_amount=row[6], status=row[7], notes=row[8]
        ))
    return result

def get_agreement_by_id(agreement_id: int) -> Optional[Agreement]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        SELECT a.id, a.supplier_id, s.name, a.agreement_number, a.date,
               a.total_amount, a.spent_amount, a.status, a.notes
        FROM agreements a
        LEFT JOIN suppliers s ON a.supplier_id = s.id
        WHERE a.id=?
    """, (agreement_id,))
    row = c.fetchone()
    conn.close()
    if row:
        date_obj = datetime.date.fromisoformat(row[4])
        return Agreement(
            id=row[0], supplier_id=row[1], supplier_name=row[2] or "",
            agreement_number=row[3], date=date_obj, total_amount=row[5],
            spent_amount=row[6], status=row[7], notes=row[8]
        )
    return None

def add_agreement(agreement: Agreement) -> int:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT INTO agreements (supplier_id, agreement_number, date, total_amount, spent_amount, status, notes)
        VALUES (?,?,?,?,?,?,?)
    """, (agreement.supplier_id, agreement.agreement_number, agreement.date.isoformat(),
          agreement.total_amount, agreement.spent_amount, agreement.status, agreement.notes))
    conn.commit()
    aid = c.lastrowid
    conn.close()
    return aid

def update_agreement(agreement: Agreement):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        UPDATE agreements SET agreement_number=?, date=?, total_amount=?, spent_amount=?, status=?, notes=?
        WHERE id=?
    """, (agreement.agreement_number, agreement.date.isoformat(),
          agreement.total_amount, agreement.spent_amount, agreement.status, agreement.notes, agreement.id))
    conn.commit()
    conn.close()

def delete_agreement(agreement_id: int):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # Удалить связанные счета, партии, документы
    c.execute("SELECT id FROM purchases WHERE agreement_id=?", (agreement_id,))
    for pid in c.fetchall():
        c.execute("DELETE FROM purchase_items WHERE purchase_id=?", (pid[0],))
        c.execute("DELETE FROM material_batches WHERE purchase_id=?", (pid[0],))
    c.execute("DELETE FROM purchases WHERE agreement_id=?", (agreement_id,))
    c.execute("DELETE FROM documents WHERE agreement_id=?", (agreement_id,))
    c.execute("DELETE FROM agreements WHERE id=?", (agreement_id,))
    conn.commit()
    conn.close()

# ---------- Материалы ----------
def get_all_materials() -> List[Material]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, name, unit, purchase_price, sale_price, stock, inventory_number, type FROM materials ORDER BY name")
    rows = c.fetchall()
    conn.close()
    return [Material(*row) for row in rows]

def get_material_by_id(material_id: int) -> Optional[Material]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, name, unit, purchase_price, sale_price, stock, inventory_number, type FROM materials WHERE id=?", (material_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return Material(*row)
    return None

def add_material(material: Material) -> int:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT INTO materials (name, unit, purchase_price, sale_price, stock, inventory_number, type) VALUES (?,?,?,?,?,?,?)",
              (material.name, material.unit, material.purchase_price, material.sale_price, material.stock, material.inventory_number, material.type))
    conn.commit()
    mid = c.lastrowid
    conn.close()
    return mid

def update_material(material: Material):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE materials SET name=?, unit=?, purchase_price=?, sale_price=?, stock=?, inventory_number=?, type=? WHERE id=?",
              (material.name, material.unit, material.purchase_price, material.sale_price, material.stock, material.inventory_number, material.type, material.id))
    conn.commit()
    conn.close()

def delete_material(material_id: int):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM materials WHERE id=?", (material_id,))
    conn.commit()
    conn.close()

# ---------- Счета и партии ----------
def get_purchases_by_agreement(agreement_id: int) -> List[Purchase]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, agreement_id, date, invoice_number, notes FROM purchases WHERE agreement_id=? ORDER BY date DESC", (agreement_id,))
    rows = c.fetchall()
    purchases = []
    for row in rows:
        date_obj = datetime.date.fromisoformat(row[2])
        p = Purchase(id=row[0], agreement_id=row[1], date=date_obj, invoice_number=row[3], notes=row[4])
        c2 = conn.cursor()
        c2.execute("SELECT id, material_id, quantity, purchase_price FROM purchase_items WHERE purchase_id=?", (p.id,))
        for pi in c2.fetchall():
            p.items.append(PurchaseItem(id=pi[0], purchase_id=p.id, material_id=pi[1], quantity=pi[2], purchase_price=pi[3]))
        purchases.append(p)
    conn.close()
    return purchases

def get_purchase_by_id(purchase_id: int) -> Optional[Purchase]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, agreement_id, date, invoice_number, notes FROM purchases WHERE id=?", (purchase_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return None
    purchase = Purchase(id=row[0], agreement_id=row[1], date=datetime.date.fromisoformat(row[2]),
                        invoice_number=row[3], notes=row[4])
    c2 = conn.cursor()
    c2.execute("SELECT id, material_id, quantity, purchase_price FROM purchase_items WHERE purchase_id=?", (purchase.id,))
    for pi in c2.fetchall():
        purchase.items.append(PurchaseItem(id=pi[0], purchase_id=purchase.id,
                                           material_id=pi[1], quantity=pi[2],
                                           purchase_price=pi[3]))
    conn.close()
    return purchase

def add_purchase(purchase: Purchase) -> int:
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("BEGIN")
        c = conn.cursor()
        c.execute("""
            INSERT INTO purchases (agreement_id, date, invoice_number, notes)
            VALUES (?,?,?,?)
        """, (purchase.agreement_id, purchase.date.isoformat(), purchase.invoice_number, purchase.notes))
        pid = c.lastrowid
        total_invoice = 0.0
        for item in purchase.items:
            c.execute("""
                INSERT INTO purchase_items (purchase_id, material_id, quantity, purchase_price)
                VALUES (?,?,?,?)
            """, (pid, item.material_id, item.quantity, item.purchase_price))
            c.execute("""
                INSERT INTO material_batches (purchase_id, material_id, quantity, purchase_price, date, batch_number)
                VALUES (?,?,?,?,?,?)
            """, (pid, item.material_id, item.quantity, item.purchase_price, purchase.date.isoformat(), ""))
            # Обновляем остаток материала
            c.execute("UPDATE materials SET stock = stock + ? WHERE id = ?", (item.quantity, item.material_id))
            total_invoice += item.quantity * item.purchase_price
        # Проверка лимита договора
        c.execute("SELECT total_amount, spent_amount FROM agreements WHERE id=?", (purchase.agreement_id,))
        total, spent = c.fetchone()
        if spent + total_invoice > total:
            raise ValueError(f"Превышение суммы договора: доступно {total - spent}, требуется {total_invoice}")
        c.execute("UPDATE agreements SET spent_amount = spent_amount + ? WHERE id=?", (total_invoice, purchase.agreement_id))
        conn.commit()
        return pid
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def update_purchase(purchase: Purchase):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("BEGIN")
        c = conn.cursor()
        # Получить старый договор и старые позиции
        c.execute("SELECT agreement_id FROM purchases WHERE id=?", (purchase.id,))
        old_agreement_id = c.fetchone()[0]
        c.execute("SELECT id, material_id, quantity, purchase_price FROM purchase_items WHERE purchase_id=?", (purchase.id,))
        old_items = c.fetchall()
        # Восстановить остатки материалов и уменьшить spent_amount старого договора
        old_total = 0.0
        for item_id, mat_id, qty, price in old_items:
            c.execute("UPDATE materials SET stock = stock - ? WHERE id = ?", (qty, mat_id))
            old_total += qty * price
        c.execute("UPDATE agreements SET spent_amount = spent_amount - ? WHERE id=?", (old_total, old_agreement_id))
        c.execute("DELETE FROM purchase_items WHERE purchase_id=?", (purchase.id,))
        c.execute("DELETE FROM material_batches WHERE purchase_id=?", (purchase.id,))

        # Обновить данные счёта
        c.execute("""
            UPDATE purchases SET agreement_id=?, date=?, invoice_number=?, notes=?
            WHERE id=?
        """, (purchase.agreement_id, purchase.date.isoformat(),
              purchase.invoice_number, purchase.notes, purchase.id))

        # Вставить новые позиции и партии
        new_total = 0.0
        for item in purchase.items:
            c.execute("""
                INSERT INTO purchase_items (purchase_id, material_id, quantity, purchase_price)
                VALUES (?,?,?,?)
            """, (purchase.id, item.material_id, item.quantity, item.purchase_price))
            c.execute("""
                INSERT INTO material_batches (purchase_id, material_id, quantity, purchase_price, date, batch_number)
                VALUES (?,?,?,?,?,?)
            """, (purchase.id, item.material_id, item.quantity, item.purchase_price,
                  purchase.date.isoformat(), ""))
            # Обновляем остаток материала
            c.execute("UPDATE materials SET stock = stock + ? WHERE id = ?", (item.quantity, item.material_id))
            new_total += item.quantity * item.purchase_price

        # Проверка лимита нового договора
        c.execute("SELECT total_amount, spent_amount FROM agreements WHERE id=?", (purchase.agreement_id,))
        total, spent = c.fetchone()
        if spent + new_total > total:
            raise ValueError(f"Превышение суммы договора: доступно {total - spent}, требуется {new_total}")

        c.execute("UPDATE agreements SET spent_amount = spent_amount + ? WHERE id=?", (new_total, purchase.agreement_id))

        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def delete_purchase(purchase_id: int):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("BEGIN")
        c = conn.cursor()
        c.execute("SELECT agreement_id FROM purchases WHERE id=?", (purchase_id,))
        agreement_id = c.fetchone()[0]
        c.execute("SELECT material_id, quantity, purchase_price FROM purchase_items WHERE purchase_id=?", (purchase_id,))
        items = c.fetchall()
        total = 0.0
        for mat_id, qty, price in items:
            c.execute("UPDATE materials SET stock = stock - ? WHERE id = ?", (qty, mat_id))
            total += qty * price
        c.execute("UPDATE agreements SET spent_amount = spent_amount - ? WHERE id=?", (total, agreement_id))
        c.execute("DELETE FROM purchase_items WHERE purchase_id=?", (purchase_id,))
        c.execute("DELETE FROM material_batches WHERE purchase_id=?", (purchase_id,))
        c.execute("DELETE FROM purchases WHERE id=?", (purchase_id,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

# ---------- Партии (выбор и списание) ----------
def get_available_batches(material_id: int) -> List[MaterialBatch]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        SELECT id, purchase_id, material_id, quantity, purchase_price, date, batch_number
        FROM material_batches
        WHERE material_id = ? AND quantity > 0
        ORDER BY date ASC
    """, (material_id,))
    rows = c.fetchall()
    conn.close()
    result = []
    for row in rows:
        date_obj = datetime.date.fromisoformat(row[5])
        result.append(MaterialBatch(id=row[0], purchase_id=row[1], material_id=row[2],
                                    quantity=row[3], purchase_price=row[4],
                                    date=date_obj, batch_number=row[6]))
    return result

def allocate_material_from_batches(material_id: int, required_quantity: float, order_material_id: int, conn: sqlite3.Connection):
    batches = get_available_batches(material_id)
    remaining = required_quantity
    c = conn.cursor()
    for batch in batches:
        if remaining <= 0:
            break
        take = min(batch.quantity, remaining)
        c.execute("UPDATE material_batches SET quantity = quantity - ? WHERE id = ?", (take, batch.id))
        c.execute("""
            INSERT INTO order_material_batches (order_material_id, material_batch_id, quantity)
            VALUES (?,?,?)
        """, (order_material_id, batch.id, take))
        remaining -= take
    if remaining > 0:
        raise ValueError(f"Недостаточно материала {material_id} (требуется {required_quantity}, доступно {required_quantity - remaining})")

def allocate_from_selected_batches(om: OrderMaterial, conn: sqlite3.Connection):
    c = conn.cursor()
    total_required = om.quantity
    allocated = 0.0
    for batch_id, qty in om.temp_batches:
        c.execute("SELECT quantity FROM material_batches WHERE id=?", (batch_id,))
        row = c.fetchone()
        if not row or row[0] < qty:
            raise ValueError(f"Недостаточно в партии {batch_id}")
        c.execute("UPDATE material_batches SET quantity = quantity - ? WHERE id=?", (qty, batch_id))
        c.execute("""
            INSERT INTO order_material_batches (order_material_id, material_batch_id, quantity)
            VALUES (?,?,?)
        """, (om.id, batch_id, qty))
        allocated += qty
    if allocated < total_required:
        remaining = total_required - allocated
        allocate_material_from_batches(om.material_id, remaining, om.id, conn)

def restore_order_materials(order_id: int, conn: sqlite3.Connection):
    c = conn.cursor()
    c.execute("""
        SELECT omb.id, omb.material_batch_id, omb.quantity
        FROM order_material_batches omb
        JOIN order_materials om ON omb.order_material_id = om.id
        JOIN work_items wi ON om.work_item_id = wi.id
        WHERE wi.order_id = ?
    """, (order_id,))
    for omb_id, batch_id, qty in c.fetchall():
        c.execute("UPDATE material_batches SET quantity = quantity + ? WHERE id = ?", (qty, batch_id))
        c.execute("DELETE FROM order_material_batches WHERE id = ?", (omb_id,))

# ---------- Заказчики ----------
def get_all_customers() -> List[Customer]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, name, phone, email, address, notes FROM customers ORDER BY name")
    rows = c.fetchall()
    conn.close()
    return [Customer(*row) for row in rows]

def get_customer_by_id(customer_id: int) -> Optional[Customer]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, name, phone, email, address, notes FROM customers WHERE id=?", (customer_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return Customer(*row)
    return None

def add_customer(customer: Customer) -> int:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT INTO customers (name, phone, email, address, notes) VALUES (?,?,?,?,?)",
              (customer.name, customer.phone, customer.email, customer.address, customer.notes))
    conn.commit()
    cid = c.lastrowid
    conn.close()
    return cid

def update_customer(customer: Customer):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE customers SET name=?, phone=?, email=?, address=?, notes=? WHERE id=?",
              (customer.name, customer.phone, customer.email, customer.address, customer.notes, customer.id))
    conn.commit()
    conn.close()

def delete_customer(customer_id: int):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM customers WHERE id=?", (customer_id,))
    conn.commit()
    conn.close()

# ---------- Заказы ----------
def get_all_orders() -> List[Order]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, customer_id, customer_name, order_number, date, status, notes FROM orders ORDER BY date DESC")
    rows = c.fetchall()
    orders = []
    for row in rows:
        order_id, cust_id, cust_name, number, date_str, status, notes = row
        date_obj = datetime.date.fromisoformat(date_str)
        order = Order(id=order_id, customer_id=cust_id, customer_name=cust_name,
                      order_number=number, date=date_obj, status=status, notes=notes)
        # work items
        c2 = conn.cursor()
        c2.execute("SELECT id, description, hours, price FROM work_items WHERE order_id=?", (order_id,))
        for wr in c2.fetchall():
            work = WorkItem(id=wr[0], order_id=order_id, description=wr[1], hours=wr[2], price=wr[3])
            # order materials
            c3 = conn.cursor()
            c3.execute("SELECT id, material_id, quantity, sale_price FROM order_materials WHERE work_item_id=?", (work.id,))
            for omr in c3.fetchall():
                om = OrderMaterial(id=omr[0], work_item_id=work.id, material_id=omr[1], quantity=omr[2], sale_price=omr[3])
                # batches
                c4 = conn.cursor()
                c4.execute("""
                    SELECT omb.id, omb.material_batch_id, omb.quantity, mb.purchase_price, mb.date
                    FROM order_material_batches omb
                    JOIN material_batches mb ON omb.material_batch_id = mb.id
                    WHERE omb.order_material_id = ?
                """, (om.id,))
                for br in c4.fetchall():
                    omb = OrderMaterialBatch(id=br[0], order_material_id=om.id, material_batch_id=br[1], quantity=br[2])
                    omb.purchase_price = br[3]
                    omb.batch_date = datetime.date.fromisoformat(br[4])
                    om.batches.append(omb)
                work.materials.append(om)
            order.work_items.append(work)
        orders.append(order)
    conn.close()
    return orders

def get_order_by_id(order_id: int) -> Optional[Order]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, customer_id, customer_name, order_number, date, status, notes FROM orders WHERE id=?", (order_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return None
    order_id, cust_id, cust_name, number, date_str, status, notes = row
    date_obj = datetime.date.fromisoformat(date_str)
    order = Order(id=order_id, customer_id=cust_id, customer_name=cust_name,
                  order_number=number, date=date_obj, status=status, notes=notes)

    c2 = conn.cursor()
    c2.execute("SELECT id, description, hours, price FROM work_items WHERE order_id=?", (order_id,))
    for wr in c2.fetchall():
        work = WorkItem(id=wr[0], order_id=order_id, description=wr[1], hours=wr[2], price=wr[3])
        c3 = conn.cursor()
        c3.execute("SELECT id, material_id, quantity, sale_price FROM order_materials WHERE work_item_id=?", (work.id,))
        for omr in c3.fetchall():
            om = OrderMaterial(id=omr[0], work_item_id=work.id, material_id=omr[1], quantity=omr[2], sale_price=omr[3])
            c4 = conn.cursor()
            c4.execute("""
                SELECT omb.id, omb.material_batch_id, omb.quantity, mb.purchase_price, mb.date
                FROM order_material_batches omb
                JOIN material_batches mb ON omb.material_batch_id = mb.id
                WHERE omb.order_material_id = ?
            """, (om.id,))
            for br in c4.fetchall():
                omb = OrderMaterialBatch(id=br[0], order_material_id=om.id, material_batch_id=br[1], quantity=br[2])
                omb.purchase_price = br[3]
                omb.batch_date = datetime.date.fromisoformat(br[4])
                om.batches.append(omb)
            work.materials.append(om)
        order.work_items.append(work)
    conn.close()
    return order

def add_order(order: Order) -> int:
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("BEGIN")
        c = conn.cursor()
        c.execute("""
            INSERT INTO orders (customer_id, customer_name, order_number, date, status, notes)
            VALUES (?,?,?,?,?,?)
        """, (order.customer_id, order.customer_name, order.order_number,
              order.date.isoformat(), order.status, order.notes))
        order_id = c.lastrowid

        for work in order.work_items:
            c.execute("""
                INSERT INTO work_items (order_id, description, hours, price)
                VALUES (?,?,?,?)
            """, (order_id, work.description, work.hours, work.price))
            work_id = c.lastrowid

            for om in work.materials:
                c.execute("""
                    INSERT INTO order_materials (work_item_id, material_id, quantity, sale_price)
                    VALUES (?,?,?,?)
                """, (work_id, om.material_id, om.quantity, om.sale_price))
                om_id = c.lastrowid
                om.id = om_id
                if hasattr(om, 'temp_batches') and om.temp_batches:
                    allocate_from_selected_batches(om, conn)
                else:
                    allocate_material_from_batches(om.material_id, om.quantity, om_id, conn)

        conn.commit()
        return order_id
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def update_order(order: Order):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("BEGIN")
        c = conn.cursor()
        restore_order_materials(order.id, conn)
        c.execute("SELECT id FROM work_items WHERE order_id=?", (order.id,))
        for wi in c.fetchall():
            c.execute("DELETE FROM order_materials WHERE work_item_id=?", (wi[0],))
        c.execute("DELETE FROM work_items WHERE order_id=?", (order.id,))
        c.execute("""
            UPDATE orders SET customer_id=?, customer_name=?, order_number=?, date=?, status=?, notes=?
            WHERE id=?
        """, (order.customer_id, order.customer_name, order.order_number,
              order.date.isoformat(), order.status, order.notes, order.id))
        for work in order.work_items:
            c.execute("""
                INSERT INTO work_items (order_id, description, hours, price)
                VALUES (?,?,?,?)
            """, (order.id, work.description, work.hours, work.price))
            work_id = c.lastrowid
            for om in work.materials:
                c.execute("""
                    INSERT INTO order_materials (work_item_id, material_id, quantity, sale_price)
                    VALUES (?,?,?,?)
                """, (work_id, om.material_id, om.quantity, om.sale_price))
                om_id = c.lastrowid
                om.id = om_id
                if hasattr(om, 'temp_batches') and om.temp_batches:
                    allocate_from_selected_batches(om, conn)
                else:
                    allocate_material_from_batches(om.material_id, om.quantity, om_id, conn)
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def delete_order(order_id: int):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("BEGIN")
        c = conn.cursor()
        restore_order_materials(order_id, conn)
        c.execute("SELECT id FROM work_items WHERE order_id=?", (order_id,))
        for wi in c.fetchall():
            c.execute("DELETE FROM order_materials WHERE work_item_id=?", (wi[0],))
        c.execute("DELETE FROM work_items WHERE order_id=?", (order_id,))
        c.execute("DELETE FROM orders WHERE id=?", (order_id,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def get_orders_by_customer_period(customer_id: int, start_date: datetime.date, end_date: datetime.date) -> List[Order]:
    all_orders = get_all_orders()
    return [o for o in all_orders if o.customer_id == customer_id and start_date <= o.date <= end_date]

# ---------- Документы ----------
def get_documents_by_supplier(supplier_id: int) -> List[Document]:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, supplier_id, agreement_id, purchase_id, document_type, file_path, file_name, uploaded_date, description FROM documents WHERE supplier_id=?", (supplier_id,))
    rows = c.fetchall()
    result = []
    for row in rows:
        uploaded = datetime.date.fromisoformat(row[7]) if row[7] else None
        result.append(Document(id=row[0], supplier_id=row[1], agreement_id=row[2], purchase_id=row[3],
                               document_type=row[4], file_path=row[5], file_name=row[6],
                               uploaded_date=uploaded, description=row[8]))
    return result

def add_document(doc: Document) -> int:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    uploaded_str = doc.uploaded_date.isoformat() if doc.uploaded_date else None
    c.execute("""
        INSERT INTO documents (supplier_id, agreement_id, purchase_id, document_type, file_path, file_name, uploaded_date, description)
        VALUES (?,?,?,?,?,?,?,?)
    """, (doc.supplier_id, doc.agreement_id, doc.purchase_id, doc.document_type, doc.file_path,
          doc.file_name, uploaded_str, doc.description))
    conn.commit()
    did = c.lastrowid
    conn.close()
    return did

def delete_document(doc_id: int):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM documents WHERE id=?", (doc_id,))
    conn.commit()
    conn.close()

def has_supplier_materials(supplier_id: int) -> bool:
    """Проверяет, есть ли у поставщика материалы (партии)."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        SELECT COUNT(*) FROM material_batches mb
        JOIN purchases p ON mb.purchase_id = p.id
        WHERE p.agreement_id IN (SELECT id FROM agreements WHERE supplier_id = ?)
    """, (supplier_id,))
    count = c.fetchone()[0]
    conn.close()
    return count > 0