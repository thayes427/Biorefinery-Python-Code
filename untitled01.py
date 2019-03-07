from multiprocessing import Process, Manager
from textwrap import wrap
import Illuminate_Simulations

def f(d, l):
    d[1] = '1'
    d['2'] = 2
    d[0.25] = None
    l.reverse()


class obb:
    def __init__(self):
        self.manager=Manager()
        print('ss')


if __name__ == '__main__':
    manager = obb()
    manager = manager.manager

    d = manager.dict()
    l = manager.list(range(10))

    p = Process(target=f, args=(d, l))
    p.start()
    p.join()

