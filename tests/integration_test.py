from autocore import *

# ============================================================
#  LOG SETUP TEST
# ============================================================
print("Log setup test ")
print(30 * '=')
log_setup("integration_test")
print("Log setup working")
print("This should appear in logs folder")
print(30 * '=')
# ============================================================
#  IMPORT TEST
# ============================================================
print(" Import test done")
print(30 * '=')

# ============================================================
#  DATE & TIME TESTS
# ============================================================
print(" Date & Time tests ")
print(30 * '=')
print("Year       :", year())
print("Month      :", month())
print("Date       :", date())
print("Day        :", day())
print("Hour       :", hour())
print("Minute     :", minute())
print("Second     :", second())
print(30 * '=')

# ============================================================
#  STRING & DATA TESTS
# ============================================================
print("String & Data tests ")
print(30 * '=')
print("find_str   :", find_str("Hello World this is a test", "Hello ", " this"))  # expected: World
data = {'user': {'id': 1, 'name': 'Ash'}, 'admin': {'id': 2, 'name': 'Bob'}}
print("find_key id    :", find_key(data, 'id'))  # expected: [1, 2]
print("find_key name  :", find_key(data, 'name'))  # expected: ['Ash', 'Bob']
print(30 * '=')

# ============================================================
#  FILE READ TESTS
# ============================================================
print(" File read tests")
print(30 * '=')
with open('test.txt', 'w') as f:
    f.write('Hello AutoCore')
print("Read txt            :", read('test.txt'))

with open('test.json', 'w') as f:
    json.dump({'name': 'Ash', 'tool': 'AutoCore'}, f)
print("Read json           :", read('test.json'))

with open('test.csv', 'w') as f:
    f.write('name,age\nChris,25\nBob,30')
print("Read csv            :", read('test.csv'))

print("Read auto detect    :", read('test'))
print(30 * '=')

# ============================================================
#  CSV TO XLSX TEST
# ============================================================
print(" CSV to XLSX test ")
print(30 * '=')
with open('convert_test.csv', 'w') as f:
    f.write('name,age\nChris,25\nBob,30')
print("CSV to XLSX         :", csv_to_xlsx('convert_test.csv', delete_csv=False))
print(30 * '=')

# ============================================================
#  WAIT TESTS
# ============================================================
print(" Wait tests ")
print(30 * '=')
print("Wait countdown      :", wait(3))
print("Wait silent         :", wait(3, countdown=False))
print(30 * '=')



# ============================================================
#  SAY TEST (requires speakers)
# ============================================================
print(" Say test ")
print(30 * '=')
say("Testing volume 50%", volume=0.5)
say("AutoCore voice test done")
print(30 * '=')

# ============================================================
#  OCR TEST
# ============================================================
print(" OCR test ")
print(30 * '=')
text = read()
print("OCR text   :", text)
print("OCR type   :", type(text))
print(30 * '=')

# ============================================================
#  SCREENSHOT TESTS
# ============================================================
print("Screenshot tests ")
print(30 * '=')
screenshot()
screenshot('test_full.png')
screenshot(0, 0, 500, 500)
screenshot(0, 0, 500, 300, 'test_region_with_custom_name.png')
print(30 * '=')

# ============================================================
#  BROWSER TESTS
# ============================================================
print(" Browser launch tests")
print(30 * '=')
url = "https://youtube.com"
dr = browser(url)
print("Browser type        :", type(dr))
dr2 = browser(url, headless=True)
print("Browser headless    :", type(dr2))
screenshot(dr2, "headless_browser.png")
print("Screenshot of headless browser taken.")
print(30 * '=')

# ============================================================
#  BROWSER OCR READ TEST
# ============================================================
print(" Browser read test using OCR")
print(30 * '=')
text = read(dr)
print("Browser read text for visible browser  :", text)
print(20 * '-')
text = read(dr2)
print("Browser read text for headless browser  :", text)
print(30 * '=')

# ============================================================
#  BROWSER TEXT READ TEST
# ============================================================
print("Browser read test using copy")
print(30 * '=')
text = copy(dr)
print("Browser read text for visible browser  :", text)
print(20 * '-')
text = copy(dr2)
print("Browser read text for headless browser  :", text)
print(30 * '=')

# ============================================================
#  CLICK TESTS
# ============================================================
print(" Selenium click tests ")
print(30 * '=')
print("Click selenium      :", click(dr, 'text', 'Music'))
print("Click selenium      :", click(dr2, 'text', 'Music'))
time.sleep(3)
screenshot(dr2, "headless browser after click test")
print(30 * '=')

# ============================================================
#  WRITE & PRESS TESTS
# ============================================================
print(" Write & Press tests ")
print(30 * '=')
print("Write result        :", write(dr, 'name', "search_query", "New song"))
print("Write result        :", write(dr2, 'name', "search_query", "New song"))
print("Press result        :", press(dr, 'enter'))
print("Press result        :", press(dr2, 'enter'))
time.sleep(3)
screenshot(dr2, "headless browser after Write & Press test")
print(30 * '=')

# ============================================================
#  ERASE TEST
# ============================================================
print(" Erase test ")
print(30 * '=')
erase(dr, 'name', "search_query")
erase(dr2, 'name', "search_query")
time.sleep(3)
screenshot(dr2, "headless browser after erase test")
print(30 * '=')

# ============================================================
#  SCROLL TESTS
# ============================================================
print(" Scroll tests ")
print(30 * '=')
print("Scroll down 3       :", scroll(dr, 'down', 3))
print("Scroll up 2         :", scroll(dr, 'up', 2))
print("Scroll top       :", scroll(dr, 'top'))
print("Scroll bottom          :", scroll(dr, 'bottom'))

print("Scroll headless down 3       :", scroll(dr2, 'down', 3))
print("Scroll headless up 2         :", scroll(dr2, 'up', 2))
print("Scroll headless to top       :", scroll(dr2, 'top'))
print("Scroll headless to bottom          :", scroll(dr2, 'bottom'))
time.sleep(3)
screenshot(dr2, "headless browser after scroll test")
print(30 * '=')

# ============================================================
# ZOOM TEST
# ============================================================
print("Zoom test ")
print(30 * '=')
print("Zoom in 3           :", zoom(dr, 3))
print("Zoom out 3          :", zoom(dr, -3))
print("Zoom reset          :", zoom(dr, 100))

print("Zoom in headless 3     :", zoom(dr2, 3))
print("Zoom out headless 3    :", zoom(dr2, -3))
print("Zoom reset headless    :", zoom(dr2, 100))
time.sleep(3)
screenshot(dr2, "headless browser after zoom test")
print(30 * '=')


# ============================================================
#  WINDOW TESTS
# ============================================================
print(" Window tests ")
print(30 * '=')
print("Window list         :", window())
print("Window title        :", window('title'))
print("Window focus        :", window('focus', 'YouTube'))
print("Window resize       :", window('resize', 'YouTube', 800, 600))
print("Window move         :", window('move', 'Google Chrome', 100, 100))
print("Window minimize     :", window('minimize', 'Google Chrome'))
print("Window maximize     :", window('maximize', 'Google Chrome'))
print(30 * '=')

# ============================================================
#  WAIT DOWNLOAD TEST
# ============================================================
print(" Wait download test ")
print(30 * '=')
dr.get('https://docs.python.org/3/download.html')
click(dr, 'xpath', "(//a[contains(text(),'Download')])[4]")
print("Wait download       :", wait_download(5))

dr2.get('https://docs.python.org/3/download.html')
click(dr2, 'xpath', "(//a[contains(text(),'Download')])[4]")
print("Wait download for headless      :", wait_download(5))
print(30 * '=')

# ============================================================
# FIND BROWSER SELENIUM TEST
# ============================================================
print(" Find browser test ")
print(30 * '=')
click(dr, 'text', "modules")
print("Find browser selenium :", find_browser(dr, 'marshal'))
wait(2)

click(dr2, 'text', "modules")
print("Find headless browser selenium :", find_browser(dr2, 'marshal'))
screenshot(dr2, "headless browser after find browser test")
print(30 * '=')

# ============================================================
# FIND BROWSER UI TEST
# ============================================================
print(" Find browser ui test ")
print(30 * '=')
window('focus', 'Python Module Index')
print("Find browser using ui :", find_browser('argparse'))
wait(2)
print(30 * '=')

# ============================================================
#  DROPDOWN TEST
# ============================================================
print(" Dropdown test ")
print(30*'=')
dr.get("https://www.globalsqa.com/demo-site/select-dropdown-menu/")
print("Dropdown result     :", dropdown_select(dr, 'xpath', "(//select)[1]" , 'United States'))
wait(2)

dr2.get("https://www.globalsqa.com/demo-site/select-dropdown-menu/")
print("Dropdown result     :", dropdown_select(dr2, 'xpath', "(//select)[1]" , 'United States'))
screenshot(dr2, "headless browser after dropdown test")
print(30*'=')

# ============================================================
#  DRAG TEST
# ============================================================
print(" Drag test ")
print(30 * '=')
print("Drag result         :", drag(200, 200, 400, 400))
print(30 * '=')

# ============================================================
# PRESS COMBINATIONS TEST
# ============================================================
print(" Press combinations test ")
print(30 * '=')
print("Press ctrl+a        :", press('ctrl', 'a'))
print("Press ctrl+c        :", press('ctrl', 'c'))
print("Result after ctrl+a and ctrl+c :", copy('clipboard'))

print("Press tab 3         :", press('tab', 3))
print("Press shift+tab 3   :", press('tab', -3))
print(30 * '=')

# ============================================================
# CLEANUP
# ============================================================
# for f in ['test.txt', 'test.json', 'test.csv', 'convert_test.csv', 'convert_test.xlsx']:
#     if os.path.exists(f):
#         os.remove(f)
#         print(f"Deleted             : {f}")
# print("25. Cleanup done")
dr.quit()
dr2.quit()
print(30 * '=')
print("All integration tests completed")
print(30 * '=')