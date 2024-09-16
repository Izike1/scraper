from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

options = webdriver.ChromeOptions()
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

try:
    driver.get('https://www.oddsportal.com/football/england/premier-league/results/')
    time.sleep(4)

    for _ in range(2):
        driver.execute_script("window.scrollBy(0, 1000);")  
        time.sleep(2)  

    for _ in range(2):
        driver.execute_script("window.scrollBy(0, -1000);")  
        time.sleep(2)  

    games = driver.find_elements(By.CSS_SELECTOR, "a[href*='/football/england/premier-league/']")
    if not games:
        raise Exception("Ссылки на игры не найдены, возможно, контент не подгружен.")

    games_set = set([el.get_attribute('href') for el in games])

    game_data = []
    for game in list(games_set):
        url = game + '#over-under;2'
        driver.get(url)
        time.sleep(2)
        try:
            date_element = driver.find_elements(By.CSS_SELECTOR, "div[class^='text-gray-dark font-main item-center flex gap-1 text-xs font-normal']")
            game_date, game_time = date_element[0].text.split(',')[1:]
        except (IndexError, ValueError):
            game_date, game_time = "N/A", "N/A"

        try:
            teams = driver.find_elements(By.CSS_SELECTOR, "div[class^='max-mm:flex-col min-sm:w-full min-mm:items-center text-black-main font-secondary relative flex w-auto justify-center gap-2 px-[12px] pb-3 pt-5 text-[22px] font-semibold']")
            team_text = teams[0].text.split('\n')
            home = team_text[0] if len(team_text) > 0 else "N/A"
            away = team_text[-2] if len(team_text) > 1 else "N/A"
        except (IndexError, ValueError):
            home, away = "N/A", "N/A"

        try:
            result = driver.find_elements(By.CSS_SELECTOR, "div[class^='min-sm:items-center border-black-borders mt-5 flex flex-wrap gap-4 border-t px-[10px] pt-5 text-xs max-sm:flex-col max-sm:gap-2']")
            full_time = result[0].text
        except (IndexError, ValueError):
            full_time = "N/A"

        coef_list = []
        over_under_value_start = 2
        over_under_value_end = 4
        start_value = 0.5
        end_value = 6.5
        over_under_value_start_current = over_under_value_start
        current_value = start_value
        over_under_step = 0
        while over_under_value_start_current <= over_under_value_end:
            tail = f'#over-under;{over_under_value_start_current};{current_value};0'
            coef_url = game + tail
            driver.get(coef_url)
            driver.execute_script("window.scrollTo(0, 1000);")
            time.sleep(2)

            while start_value <= end_value:
                over_under = driver.find_elements(By.CSS_SELECTOR, "div[class^='border-black-borders hover:bg-gray-light flex h-9 cursor-pointer border-b border-l border-r text-xs']")
                if over_under:
                    over_under_list = list(over_under)
                    if len(over_under_list) > 15:
                        actions = ActionChains(driver)
                        actions.click(over_under[over_under_step]).perform()
                time.sleep(2)
                sportsbook = driver.find_elements(By.CSS_SELECTOR, "p[class^='height-content max-mm:hidden pl-4']")
                pin_num = None
                for row, el in enumerate(sportsbook):
                    if el.text == "Pinnacle":
                        pin_num = row
                        break

                if pin_num is not None:
                    for _ in range(2):
                        coef = sportsbook[pin_num + 1]
                        # actions = ActionChains(driver)
                        # actions.move_to_element(coef).perform()
                        # time.sleep(3)

                        all_odds = driver.find_elements(By.CSS_SELECTOR, "div[class^='flex flex-row items-center gap-[3px]']")
                        for odd in all_odds:
                            odds_text = odd.text
                            coef_list.append(odds_text)
                        
                        time.sleep(3)
                current_value += 0.5
                over_under_step += 1

            over_under_value_start_current += 1
            time.sleep(2)
            next_page = driver.find_elements(By.CSS_SELECTOR, ".flex.gap-1.py-2.text-xs.tab-wrapper div")
            if next_page:
                next_page_list = list(next_page)
                if len(next_page_list) > over_under_value_start - 2:
                    actions = ActionChains(driver)
                    actions.click(next_page_list[over_under_value_start - 2]).perform()
            time.sleep(2)
        
        print('Дата:', game_date)
        print('Время:', game_time) 
        print('Команды:', home, 'vs', away)
        print('Счёт:', full_time)
        print('Коэффициенты:', coef_list)

        # Добавление данных
        game_data.append(
            (game_date, game_time, home, away, full_time, game, *coef_list)
        )

    # Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Дата", "Время", "Домашняя команда", "Гостевая команда", "Результат", "Ссылка на игру", "Коэффициенты"])
    for data in game_data:
        ws.append(data)
    wb.save("game_data_over_under.xlsx")

finally:
    print('ok')