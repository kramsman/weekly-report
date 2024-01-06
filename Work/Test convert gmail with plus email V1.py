import re

s = 'Kramsman+test@Gmail.com'
s2 = re.sub('\+[^>]+@gmail.com', '@gmail.com', s, flags=re.IGNORECASE)

print(f"{s=}, {s2=}")
