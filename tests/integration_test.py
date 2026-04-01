from autocore import *
import os
import json

# ============================================================
# 1. IMPORT TEST
# ============================================================
print("1. Import test done")

# ============================================================
# 2. DATE & TIME TESTS
# ============================================================
print("Year       :", year())
print("Month      :", month())
print("Date       :", date())
print("Day        :", day())
print("Hour       :", hour())
print("Minute     :", minute())
print("Second     :", second())
print("2. Date & Time tests done")

# ============================================================
# 3. STRING & DATA TESTS
# ============================================================
print("find_str   :", find_str("Hello World this is a test", "Hello ", " this"))  # expected: World

data = {'user': {'id': 1, 'name': 'Ash'}, 'admin': {'id': 2, 'name': 'Bob'}}
print("find_key id    :", find_key(data, 'id'))    # expected: [1, 2]
print("find_key name  :", find_key(data, 'name'))  # expected: ['Ash', 'Bob']
print("3. String & Data tests done")

# ============================================================
# 4. OCR TEST
# ============================================================
text = read()
print("OCR text   :", text)
print("OCR type   :", type(text))
print("4. OCR test done")

# ============================================================
# 5. SCREENSHOT TESTS
# ============================================================
screenshot()
screenshot('test_full.png')
screenshot(0, 0, 500, 300)
screenshot(0, 0, 500, 300, 'test_region.png')
print("5. Screenshot tests done")

# ============================================================
# 6. BROWSER TESTS
# ============================================================
dr = browser('https://google.com')
print("Browser type        :", type(dr))
dr2 = browser('https://google.com', headless=True)
print("Browser headless    :", type(dr2))
print("6. Browser tests done")

# ============================================================
# 7. BROWSER READ TEST
# ============================================================
dr = browser('https://google.com')
text = read(dr)
print("Browser read text   :", text)
print("7. Browser read test done")

# ============================================================
# 8. CLICK TESTS
# ============================================================
dr = browser('https://google.com')
print("Click selenium      :", click(dr, 'name', 'q'))
print("Click coordinates   :", click(500, 500))
print("Click OCR           :", click('Search'))
print("8. Click tests done")

# ============================================================
# 9. WRITE & PRESS TESTS
# ============================================================
dr = browser('https://google.com')
click(dr, 'name', 'q')
print("Write result        :", write(dr, 'name', 'q', 'autocore python'))
print("Press result        :", press(dr, 'enter'))
print("9. Write & Press tests done")

# ============================================================
# 10. COPY TESTS
# ============================================================
print("Copy active window  :", copy())
print("Copy clipboard      :", copy('clipboard'))
dr = browser('https://google.com')
print("Copy selenium       :", copy(dr, 'name', 'q'))
print("10. Copy tests done")

# ============================================================
# 11. SCROLL TESTS
# ============================================================
dr = browser('https://google.com')
print("Scroll down 3       :", scroll(dr, 'down', 3))
print("Scroll up 2         :", scroll(dr, 'up', 2))
print("Scroll bottom       :", scroll(dr, 'bottom'))
print("Scroll top          :", scroll(dr, 'top'))
print("11. Scroll tests done")

# ============================================================
# 12. WINDOW TESTS
# ============================================================
print("Window list         :", window())
print("Window title        :", window('title'))
print("Window focus        :", window('focus', 'Google Chrome'))
print("Window resize       :", window('resize', 'Google Chrome', 800, 600))
print("Window move         :", window('move', 'Google Chrome', 100, 100))
print("Window minimize     :", window('minimize', 'Google Chrome'))
print("Window maximize     :", window('maximize', 'Google Chrome'))
print("12. Window tests done")

# ============================================================
# 13. FILE READ TESTS
# ============================================================
with open('test.txt', 'w') as f:
    f.write('Hello AutoCore')
print("Read txt            :", read('test.txt'))

with open('test.json', 'w') as f:
    json.dump({'name': 'Ash', 'tool': 'AutoCore'}, f)
print("Read json           :", read('test.json'))

with open('test.csv', 'w') as f:
    f.write('name,age\nAsh,25\nBob,30')
print("Read csv            :", read('test.csv'))

print("Read auto detect    :", read('test'))
print("13. File read tests done")

# ============================================================
# 14. CSV TO XLSX TEST
# ============================================================
with open('convert_test.csv', 'w') as f:
    f.write('name,age\nAsh,25\nBob,30')
print("CSV to XLSX         :", csv_to_xlsx('convert_test.csv', delete_csv=False))
print("14. CSV to XLSX test done")

# ============================================================
# 15. WAIT TESTS
# ============================================================
print("Wait countdown      :", wait(3))
print("Wait silent         :", wait(3, countdown=False))
print("15. Wait tests done")

# ============================================================
# 16. LOG SETUP TEST
# ============================================================
log_setup("integration_test")
print("Log setup working")
print("This should appear in logs folder")
print("16. Log setup test done")

# ============================================================
# 17. DRAG TEST
# ============================================================
print("Drag result         :", drag(200, 200, 400, 400))
print("17. Drag test done")

# ============================================================
# 18. ERASE TEST
# ============================================================
dr = browser('https://google.com')
click(dr, 'name', 'q')
write(dr, 'name', 'q', 'test text')
print("Erase result        :", erase(dr, 'name', 'q'))
print("18. Erase test done")

# ============================================================
# 19. DROPDOWN TEST
# ============================================================
dr = browser('https://www.w3schools.com/tags/tryit.asp?filename=tryhtml_select')
print("Dropdown result     :", dropdown_select(dr, 'id', 'cars', 'Saab'))
print("19. Dropdown test done")

# ============================================================
# 20. ZOOM TEST
# ============================================================
dr = browser('https://google.com')
print("Zoom in 3           :", zoom(dr, 3))
print("Zoom out 3          :", zoom(dr, -3))
print("Zoom reset          :", zoom(dr, 100))
print("20. Zoom test done")

# ============================================================
# 21. SAY TEST (requires speakers)
# ============================================================
say("AutoCore integration test complete")
say("Testing volume", volume=0.5)
print("21. Say test done")

# ============================================================
# 22. WAIT DOWNLOAD TEST
# ============================================================
dr = browser('https://www.w3schools.com')
click(dr, 'partial', 'Download')
print("Wait download       :", wait_download(5))
print("22. Wait download test done")

# ============================================================
# 23. FIND BROWSER TEST
# ============================================================
dr = browser('https://google.com')
print("Find browser selenium :", find_browser(dr, 'Google'))
print("Find browser pyautogui:", find_browser('Google'))
print("23. Find browser test done")

# ============================================================
# 24. PRESS COMBINATIONS TEST
# ============================================================
print("Press ctrl+a        :", press('ctrl', 'a'))
print("Press ctrl+c        :", press('ctrl', 'c'))
print("Press tab 3         :", press('tab', 3))
print("Press shift+tab 3   :", press('tab', -3))
print("24. Press combinations test done")

# ============================================================
# 25. CLEANUP
# ============================================================
for f in ['test.txt', 'test.json', 'test.csv', 'convert_test.csv', 'convert_test.xlsx']:
    if os.path.exists(f):
        os.remove(f)
        print(f"Deleted             : {f}")
print("25. Cleanup done")
print("="*50)
print("All integration tests completed")
print("="*50)