from Hello import Calc

class Inherit(Calc):
    num2=200

    def __init__(self):
        Calc.__init__(self,2,10)

    def jai(self):
        return self.num2 + self.num
obj=Inherit()
print(obj.jai())