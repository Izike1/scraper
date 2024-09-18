from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import re
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

options = webdriver.ChromeOptions()
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

try:
    # Основная ссылка
    driver.get('https://www.oddsportal.com/football/england/premier-league/results/')
    time.sleep(4)

    for _ in range(2):
        driver.execute_script("window.scrollBy(0, 1000);")  
        time.sleep(2)  

    for _ in range(2):
        driver.execute_script("window.scrollBy(0, -1000);")  
        time.sleep(2)  

    # Также меняем href
    games = driver.find_elements(By.CSS_SELECTOR, "a[href*='/football/england/premier-league/']")
    if not games:
        raise Exception("Ссылки на игры не найдены, возможно, контент не подгружен.")

    games_set = set([el.get_attribute('href') for el in games])

    game_data = []
    # Если нужны не все игры, то добавляем [:1] к for, например for game in list(games_set)[:1]:, 1 это количество игр, то бишь можно 2,3,4 и тд. 
    for game in list(games_set)[:1]:
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
        start_value = 2
        end_value = 4
        current_value = start_value
        visited_urls = set()

        while current_value <= end_value:
            over_under_value_start = 0.50
            over_under_value_end = 8.50
            current_over_under_value = over_under_value_start

            while current_over_under_value <= over_under_value_end:
                formatted_over_under_value = "{:.2f}".format(current_over_under_value)
                tail = f'#over-under;{current_value};{formatted_over_under_value};0'
                coef_url = game + tail

                
                if coef_url in visited_urls:
                    print(f"Страница {coef_url} уже обработана, пропускаем.")
                    current_over_under_value += 0.25
                    continue

                print(f"Переход на страницу: {coef_url}")

                driver.get(coef_url)
                visited_urls.add(coef_url)

                next_page_over_under = driver.find_elements(By.CSS_SELECTOR, "div[class^='border-black-borders hover:bg-gray-light flex h-9 cursor-pointer border-b border-l border-r text-xs']")
                if next_page_over_under:
                    next_page_over_under_list = list(next_page_over_under)
                    over_under_index = int((current_over_under_value - over_under_value_start) / 0.25)

                    if over_under_index < len(next_page_over_under_list):
                        next_page_over_under_button = next_page_over_under_list[over_under_index]
                        actions = ActionChains(driver)
                        actions.click(next_page_over_under_button).perform()
                        print(f"Переход на следующую страницу over/under для значения {current_over_under_value}.")
                        time.sleep(4)
                    else:
                        print(f"Нет доступных страниц для over/under для значения {current_over_under_value}")
                else:
                    print(f"Навигационные кнопки для перехода на страницу over/under не найдены.")

                driver.refresh()
                time.sleep(3)  

                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "p[class^='height-content max-mm:hidden pl-4']"))
                    )
                    print(f"Элементы на странице {coef_url} загружены.")
                except Exception as e:
                    print(f"Ошибка загрузки страницы {coef_url}: {e}")
                    current_over_under_value += 0.25
                    continue

                driver.execute_script("window.scrollTo(0, 1000);")
                time.sleep(4)

                sportsbook = driver.find_elements(By.CSS_SELECTOR, "p[class^='height-content max-mm:hidden pl-4']")

                if not sportsbook:
                    print(f"Коэффициенты на странице {current_value} не найдены.")
                else:
                    try:
                        coef = sportsbook[0]
                        actions = ActionChains(driver)
                        actions.move_to_element(coef).perform()
                        time.sleep(2)

                        all_odds = driver.find_elements(By.CSS_SELECTOR, "div[class^='flex flex-row items-center gap-[3px]']")
                        if not all_odds:
                            print(f"Коэффициенты на странице {current_value} не найдены.")
                        else:
                            for odd in all_odds:
                                odds_text = odd.text
                                if odds_text:
                                    coef_list.append(odds_text)
                            print(f"Коэффициенты на странице {current_value} собраны: {coef_list}")
                    except Exception as e:
                        print(f"Ошибка при сборе коэффициентов на странице {current_value}: {e}")

                current_over_under_value += 0.25
                time.sleep(2)

            next_page = driver.find_elements(By.CSS_SELECTOR, ".flex.gap-1.py-2.text-xs.tab-wrapper div")
            if next_page:
                next_page_list = list(next_page)
                if current_value - start_value < len(next_page_list):
                    next_page_button = next_page_list[current_value - start_value]
                    actions = ActionChains(driver)
                    actions.click(next_page_button).perform()
                    print(f"Переход на следующую страницу {current_value + 1}.")
                    time.sleep(4)
                else:
                    print(f"Нет доступных страниц для current_value={current_value}")
            else:
                print(f"Навигационные кнопки для перехода на страницу {current_value + 1} не найдены.")

            current_value += 1
            time.sleep(2)

        
        print('Дата:', game_date)
        print('Время:', game_time) 
        print('Команды:', home, 'vs', away)
        print('Счёт:', full_time)
        print('Коэффициенты:', coef_list)

        game_data.append(
            (game_date, game_time, home, away, full_time, game, *coef_list)
        )

    # Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Дата", "Время", "Домашняя команда", "Гостевая команда", "Результат", "Ссылка на игру", "Коэффициенты"])
    for data in game_data:
        ws.append(data)

    safe_url = re.sub(r'[\/:*?"<>|]', '_', url)
    file_name = f"game_data_{safe_url}.xlsx"

    wb.save(file_name)

finally:
    driver.quit()