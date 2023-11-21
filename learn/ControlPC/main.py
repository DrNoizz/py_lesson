import pyautogui as pg
import time

#print(pg.position())
#pg.rightClick(798, 1063) # Win
#pg.leftClick(1764, 1058) # Yandex Browser

# pg.leftClick(1376, 1056) # Core Temp
# pg.moveTo(1170, 700, 0.5)
# time.sleep(5)
# pg.leftClick(1170, 699)

# Выключение ПК
pg.rightClick(798, 1063)
pg.moveTo(786, 972, 0.5)
pg.leftClick()
time.sleep(1)
pg.moveTo(1018, 937, 0.5)
time.sleep(1)
pg.doubleClick()
