import datetime
from typing import List, Optional

class Supplier:
    def __init__(self, id: int = None, name: str = "", phone: str = "",
                 email: str = "", address: str = "", notes: str = ""):
        self.id = id
        self.name = name
        self.phone = phone
        self.email = email
        self.address = address
        self.notes = notes

class Agreement:
    def __init__(self, id: int = None, supplier_id: int = 0, supplier_name: str = "",
                 agreement_number: str = "", date: datetime.date = None,
                 total_amount: float = 0.0, spent_amount: float = 0.0,
                 status: str = "активен", notes: str = ""):
        self.id = id
        self.supplier_id = supplier_id
        self.supplier_name = supplier_name
        self.agreement_number = agreement_number
        self.date = date or datetime.date.today()
        self.total_amount = total_amount
        self.spent_amount = spent_amount
        self.status = status
        self.notes = notes

    @property
    def remaining_amount(self) -> float:
        return self.total_amount - self.spent_amount

class Material:
    def __init__(self, id: int = None, name: str = "", unit: str = "",
                 purchase_price: float = 0.0, sale_price: float = 0.0,
                 stock: float = 0.0, inventory_number: str = "",
                 type: str = "материал"):  # тип: материал, гсо, расходник
        self.id = id
        self.name = name
        self.unit = unit
        self.purchase_price = purchase_price
        self.sale_price = sale_price
        self.stock = stock
        self.inventory_number = inventory_number
        self.type = type

class Purchase:
    def __init__(self, id: int = None, agreement_id: int = 0,
                 date: datetime.date = None, invoice_number: str = "", notes: str = ""):
        self.id = id
        self.agreement_id = agreement_id
        self.date = date or datetime.date.today()
        self.invoice_number = invoice_number
        self.notes = notes
        self.items: List['PurchaseItem'] = []

class PurchaseItem:
    def __init__(self, id: int = None, purchase_id: int = 0, material_id: int = 0,
                 quantity: float = 0.0, purchase_price: float = 0.0):
        self.id = id
        self.purchase_id = purchase_id
        self.material_id = material_id
        self.quantity = quantity
        self.purchase_price = purchase_price

class MaterialBatch:
    def __init__(self, id: int = None, purchase_id: int = 0, material_id: int = 0,
                 quantity: float = 0.0, purchase_price: float = 0.0,
                 date: datetime.date = None, batch_number: str = ""):
        self.id = id
        self.purchase_id = purchase_id
        self.material_id = material_id
        self.quantity = quantity
        self.purchase_price = purchase_price
        self.date = date or datetime.date.today()
        self.batch_number = batch_number

class Customer:
    def __init__(self, id: int = None, name: str = "", phone: str = "",
                 email: str = "", address: str = "", notes: str = ""):
        self.id = id
        self.name = name
        self.phone = phone
        self.email = email
        self.address = address
        self.notes = notes

class Order:
    def __init__(self, id: int = None, customer_id: int = 0, customer_name: str = "",
                 order_number: str = "", date: datetime.date = None,
                 status: str = "", notes: str = ""):
        self.id = id
        self.customer_id = customer_id
        self.customer_name = customer_name
        self.order_number = order_number
        self.date = date or datetime.date.today()
        self.status = status
        self.notes = notes
        self.work_items: List['WorkItem'] = []

class WorkItem:
    def __init__(self, id: int = None, order_id: int = 0,
                 description: str = "", hours: float = 0.0, price: float = 0.0):
        self.id = id
        self.order_id = order_id
        self.description = description
        self.hours = hours
        self.price = price
        self.materials: List['OrderMaterial'] = []

class OrderMaterial:
    def __init__(self, id: int = None, work_item_id: int = 0, material_id: int = 0,
                 quantity: float = 0.0, sale_price: float = 0.0):
        self.id = id
        self.work_item_id = work_item_id
        self.material_id = material_id
        self.quantity = quantity
        self.sale_price = sale_price
        self.batches: List['OrderMaterialBatch'] = []
        self.temp_batches = None  # временное поле для выбранных партий

class OrderMaterialBatch:
    def __init__(self, id: int = None, order_material_id: int = 0,
                 material_batch_id: int = 0, quantity: float = 0.0):
        self.id = id
        self.order_material_id = order_material_id
        self.material_batch_id = material_batch_id
        self.quantity = quantity

class Document:
    def __init__(self, id: int = None, supplier_id: int = 0, agreement_id: int = 0,
                 purchase_id: int = 0, document_type: str = "", file_path: str = "",
                 file_name: str = "", uploaded_date: datetime.date = None,
                 description: str = ""):
        self.id = id
        self.supplier_id = supplier_id
        self.agreement_id = agreement_id
        self.purchase_id = purchase_id
        self.document_type = document_type
        self.file_path = file_path
        self.file_name = file_name
        self.uploaded_date = uploaded_date or datetime.date.today()
        self.description = description