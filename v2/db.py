import json

class ProductDB(object):
    def __init__(self):
        self.products = []
        self.load_products_data()
    
    # return all product
    def all(self):
        return self.products
    
    # Insert Data
    def insert(self, product):
        self.products.append(product)
        
    # Delete Data
    def delete_by_name(self, ProductName):  
        for product in self.products:
            if ProductName == product["Product Name"]:
                self.products.remove(product)
                break
        else:
            return False
        return True

    # Search Data
    def search_by_name(self, ProductName):
        for product in self.products:
            if ProductName == product["Product Name"]:
                return product  # Product Info
        else:
            return False

    # Change Data
    def update(self, pdt):  
        ProductName = pdt["Product Name"]
        for product in self.products:
            if ProductName == product["Product Name"]:
                product.update(pdt)
                return True
        else:
            return False
        
        
    def query_by_name(self, ProductName):
        for product in self.products:
            if ProductName == product["Product Name"]:
                return True
        else:
            return False        
        
    # Get names of products        
    def name_base(self):
        database = []
        for product in self.products:
            database.append(product["Product Name"])
        return database    




    # Load Data
    def load_products_data(self):
        with open("products.json", "r", encoding="utf-8") as f:
            text = f.read()
        if text:
            self.products = json.loads(text)

    # Save Data
    def save_data(self):
        with open("products.json", 'w', encoding="utf-8") as f:
            text = json.dumps(self.products, ensure_ascii=False)
            f.write(text)

db = ProductDB()
