import json
import gspread
import requests
from time import sleep
from gspread import Cell
# from background import keep_alive
import arrow
from datetime import datetime, timedelta

while True:
    try:
        gs = gspread.service_account(filename='WakaTimeCredits.json')
        sh = gs.open_by_key('1wlv26hMOgvSj0n-2aK4W98GKL-9TUm0hZOuHSY4IBsU')
        y25 = sh.worksheet("у25")
        data = sh.worksheet("Данные")
        rating = sh.worksheet("Рейтинг кодинга")
        tmpSheet = sh.worksheet("tmpSheet")
        break
    except Exception as e:
        print("Ошибка на старте, сообщение:")
        print(e)
        sleep(60)

y25table = json.load(open("y25table.json"))


def saveJson():
    json_data = json.dumps(y25table)
    with open("y25table.json", "w") as f:
        f.write(json_data)
        f.close()


def saveY25Table(key, this_week_time, language, ide, mean_per_day, project, coding_now):
    y25table[key]["this_week_time"] = this_week_time
    y25table[key]["language"] = language
    y25table[key]["ide"] = ide
    y25table[key]["mean_per_day"] = mean_per_day
    y25table[key]["project"] = project
    y25table[key]["coding_now"] = coding_now
    saveJson()


def get_wakatime_statistics(api_key, start, end):
    print(start, "\n", end)
    arr = []
    endpoint = f"https://wakatime.com/api/v1/users/current/summaries?start={start}&end={end}&api_key=" + api_key

    response = requests.get(endpoint)

    if response.status_code == 200:
        data = response.json()

        max_time = data["cumulative_total"]["digital"]
        if not max_time:
            max_time = 0
        arr.append(max_time)

        max_time_decimal = data["cumulative_total"]["decimal"]
        if not max_time_decimal:
            max_time_decimal = 0
        arr.append(str(max_time_decimal))

        lang = {}
        for i in data["data"]:
            try:
                if i["languages"][0]["name"] in lang:
                    lang[i["languages"][0]["name"]] += int(
                        i["languages"][0]["total_seconds"])
                else:
                    lang[i["languages"][0]["name"]] = int(
                        i["languages"][0]["total_seconds"])
            except:
                pass

        if len(lang) != 0:
            most_used_l = max(lang, key=lang.get)
        else:
            most_used_l = ""
        arr.append(most_used_l)

        edit = {}
        for j in data["data"]:
            try:
                if j["editors"][0]["name"] in edit:
                    edit[j["editors"][0]["name"]] += int(
                        j["editors"][0]["total_seconds"])
                else:
                    edit[j["editors"][0]["name"]] = int(j["editors"][0]["total_seconds"])
            except:
                pass

        if len(edit) != 0:
            most_used_e = max(edit, key=edit.get)
        else:
            most_used_e = ""
        arr.append(most_used_e)

        projects = {}
        for i in data["data"]:
            try:
                if i["projects"][0]["name"] in projects:
                    projects[i["projects"][0]["name"]] += int(
                        i["projects"][0]["total_seconds"])
                else:
                    projects[i["projects"][0]["name"]] = int(
                        i["projects"][0]["total_seconds"])
            except:
                pass

        if len(projects) != 0:
            most_used_p = max(projects, key=projects.get)
        else:
            most_used_p = ""
        arr.append(most_used_p)

    else:
        print("ERROR")
        return (["0:00", "0.00", "", "", ""])
    return arr


def save_mondays_and_sundays(monday, sunday):
    data.update_cell(2, 11, monday)
    data.update_cell(2, 12, sunday)


def decimal_to_digital(decimal):
    decimal_time = float(decimal)
    hours = int(decimal_time)
    minutes = int((decimal_time - hours) * 60)
    seconds = int((((decimal_time - hours) * 60) - minutes) * 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def compare_time(this, best):
    h, m, s = this.split(':')
    seconds1 = int(h) * 3600 + int(m) * 60 + int(s)
    h, m, s = best.split(':')
    seconds2 = int(h) * 3600 + int(m) * 60 + int(s)
    return seconds1 > seconds2


def top3_times():
    this_week_times = [[key, y25table[key]['this_week_time']] for key in y25table.keys() if
                       'this_week_time' in y25table[key]]
    q = sorted([[t[0], timedelta(hours=int(t[1].split(":")[0]), minutes=int(t[1].split(":")[1]),
                                 seconds=int(t[1].split(":")[2]))] for t in this_week_times], reverse=True,
               key=lambda x: x[1])
    return q[:3]


# keep_alive()

if __name__ == '__main__':
    while True:
        try:
            now = arrow.now().utcnow() + timedelta(hours=3)
            start = now.floor('week').strftime("%Y-%m-%d")
            end = now.ceil('week').strftime("%Y-%m-%d")
            all_cell_list = []
            for i in range(2, 1000):
                try:
                    key = tmpSheet.get(f"C{i}")
                    key = key[0][0]
                    if not y25table.__contains__(key):
                        y25table[key] = {}
                        y25table["users"].append(key)
                        y25table[key]["this_week_time"] = "0:00:00"
                        y25table[key]["first_place"] = 0
                        y25table[key]["second_place"] = 0
                        y25table[key]["third_place"] = 0
                        y25table[key]["best_week_time"] = "0:00:00"
                        y25table[key]["best_week"] = [None, None]

                    arr = get_wakatime_statistics(key, start, end)

                    passed_days = (datetime.now().utcnow() + timedelta(hours=3) -
                                   datetime.strptime(start, "%Y-%m-%d")).days + 1
                    try:
                        middle = decimal_to_digital(float(arr[1]) / passed_days)
                    except Exception as e:
                        print("Ошибка в получении среднего времени, сообщение:")
                        print(e)

                    arr[0] = arr[0] + ":00"
                    last_coding = y25table[key]["this_week_time"]
                    coding_now = "❌"
                    if last_coding != arr[0]:
                        coding_now = "✅"

                    saveY25Table(key, arr[0], arr[2], arr[3], middle, arr[4], coding_now)
                    cell1 = Cell(i, 4, arr[0])
                    cell2 = Cell(i, 5, arr[2])
                    cell3 = Cell(i, 6, arr[3])
                    cell4 = Cell(i, 7, middle)
                    cell5 = Cell(i, 8, arr[4])
                    cell6 = Cell(i, 9, coding_now)
                    user_cells = [cell1, cell2, cell3, cell4, cell5, cell6]
                    all_cell_list += user_cells

                except Exception as e:
                    print("Ошибка в обновлении данных, сообщение:")
                    print(e)
                    break

            y25.update_cells(all_cell_list, value_input_option="USER_ENTERED")
            endWeek = datetime.strptime(f"{end} 23:59:59", "%Y-%m-%d %H:%M:%S")
            n = datetime.now().utcnow() + timedelta(hours=3)

            if endWeek - n <= timedelta(minutes=5):
                top1, top2, top3 = top3_times()
                y25table[top1[0]]["first_place"] += 1
                y25table[top2[0]]["second_place"] += 1
                y25table[top3[0]]["third_place"] += 1
                saveJson()

                cell_top_1 = Cell(y25table["users"].index(top1[0]) + 3, 14, y25table[top1[0]]["first_place"])
                cell_top_2 = Cell(y25table["users"].index(top2[0]) + 3, 15, y25table[top2[0]]["second_place"])
                cell_top_3 = Cell(y25table["users"].index(top3[0]) + 3, 16, y25table[top3[0]]["third_place"])
                top_cells = [cell_top_1, cell_top_2, cell_top_3]
                data.update_cells(top_cells, value_input_option="USER_ENTERED")

                week = y25table["week"]
                monday, sunday = week[0], week[1]
                for j in range(2, i):
                    key = y25table["users"][j - 2]
                    this_week_time, best_week_time = y25table[key]["this_week_time"], y25table[key]["best_week_time"]

                    if compare_time(this_week_time, best_week_time):
                        y25table[key]["best_week_time"] = this_week_time
                        y25table[key]["best_week"] = week
                        saveJson()

                        time1 = Cell(j, 8, this_week_time)
                        time2 = Cell(j, 9, monday)
                        time3 = Cell(j, 10, sunday)
                        best_times = [time1, time2, time3]
                        data.update_cells(best_times, value_input_option="USER_ENTERED")

            week = y25table["week"]
            monday, sunday = week[0], week[1]
            if start != monday and end != sunday:
                y25table["week"] = [start, end]
                saveJson()
                save_mondays_and_sundays(start, end)

            print(datetime.now().utcnow() + timedelta(hours=3))

            sleep(300)
        except Exception as e:
            print("Ошибка, сообщение:")
            print(e)
            sleep(60)
