class Calc:
    num=100
    def __init__(self,a,b):
        self.first_num=a
        self.second_num=b
        print("Hello")
    def demo(self):
        return self.first_num + self.second_num + self.num
obj=Calc(21,3)
print(obj.demo())

