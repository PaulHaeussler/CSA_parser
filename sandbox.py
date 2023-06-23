

from csa_parser import CSA


csa = CSA()

c = 1
for s in csa.sections:
    print(f"{c}. {s.name}")
    c += 1