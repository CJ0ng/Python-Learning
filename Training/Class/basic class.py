class rectangle:
    def __init__(self, length, width):
        self.length = length
        self.width = width
        
    def area(self):
        area = self.length * self.width
        return area
    
    def perimeter(self):
        perimeter = (self.length*2) + (self.width*2)
        return perimeter
    
x = rectangle(4,5)
print(x.area())
print(x.perimeter())

