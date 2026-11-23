class BasicCalculator:
    def __init__(self,a,b):
        self.first_num=a
        self.second_num=b
    def addition(self):
        print("Addition:",self.first_num + self.second_num)
    def subtraction(self):
        print("Subtraction:",self.first_num - self.second_num)
    def multiplication(self):
        print("Multiplication:",self.first_num * self.second_num)
    def division(self):
        print("Division:",self.first_num / self.second_num)
obj=BasicCalculator(10,5)
obj.addition()
obj.subtraction()
obj.multiplication()
obj.division()
