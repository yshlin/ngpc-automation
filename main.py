import os
import time
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from appium import webdriver
import re
import datetime
from functools import partial
import argparse
import json
import subprocess
import urllib.parse
import locale


locale.setlocale(locale.LC_CTYPE, 'chinese')


def waitElement(by, query, driver=None):
    if None is driver:
        driver = root
    return WebDriverWait(driver, 120).until(EC.presence_of_element_located((by, query)))


def waitElements(by, query, driver=None):
    if None is driver:
        driver = root
    return WebDriverWait(driver, 120).until(EC.presence_of_all_elements_located((by, query)))


def launchRoot():
    desired_caps = {'app': 'Root', 'ms:experimental-webdriver': True, 'ms:waitForAppLaunch': 25}
    return launchApp(desired_caps)


def launchApp(desired_caps):
    dut_url = "http://127.0.0.1:4723"
    driver = webdriver.Remote(command_executor=dut_url,
                              desired_capabilities=desired_caps)
    return driver


def getDriverFromWin(win):
    win_handle1 = win.get_attribute("NativeWindowHandle")
    win_handle = format(int(win_handle1), 'x')  # convert to hex string

    # Launch new session attached to the window
    desired_caps = {"appTopLevelWindow": win_handle}
    driver = launchApp(desired_caps)
    driver.switch_to.window(win_handle)
    # print('%s: %s' % (win.get_attribute('Name'), win_handle))
    return driver


def switchWindow(driver):
    if 0 < len(driver.window_handles):
        driver.switch_to.window(driver.window_handles[0])


def matchP(ps, k):
    for p in ps:
        if k in p.get_attribute('Name'):
            return p


def openPpt(key, val, enter):
    if val is not None:
        val.click()
        if enter:
            val.send_keys(Keys.ENTER)
    pWin = waitElement(By.XPATH, f'//Window[@ClassName="PPTFrameClass" and contains(@Name, "{key}")]')
    return getDriverFromWin(pWin)


def launchPpt(key, val=None, folder=None, enter=True):
    ppt = openPpt(key, val, enter)
    ebtn = ppt.find_elements_by_xpath('//Button[@Name="????????????(E)"]')
    if 0 < len(ebtn):
        ebtn.click()
        # reopen window
        # closeWindow(ppt)
        # switchWindow(folder)
        # ppt = openPpt(key, val)
    return ppt


def listChildren(elem, i=0):
    children = elem.find_elements_by_xpath('*/*')
    if 0 < len(children):
        print('Found! %d' % len(children))
        for child in children:
            print('%s: %s (%s)' % (
                child.get_attribute('ClassName'), child.get_attribute('Name'), child.get_attribute('AutomationId')))
    return children[i]


def toggleSlideView(ppt):
    ppt.find_element_by_xpath('//TabItem[@Name="??????"]').click()
    ppt.find_element_by_xpath('//Button[@Name="???????????????"]').click()
    return ppt.find_element_by_xpath('//Pane[@Name="???????????????"]')


def copyAllSlides(ppt, pane):
    pane.send_keys(Keys.CONTROL + 'a')
    pane.send_keys(Keys.CONTROL + 'c')


def appendSlides(ppt, pane):
    pane.send_keys(Keys.END)
    pane.send_keys(Keys.CONTROL + 'v')
    pane.send_keys(Keys.CONTROL)
    pane.send_keys('k')


def insertSlides(ppt, pane, before=''):
    pane.send_keys(Keys.HOME)
    # dest = ppt.find_element_by_xpath(f'//Pane[@Name="???????????????"]//ListItem[@Name="{before}"]')
    # dest = waitElement(By.XPATH, f'//ListItem[@Name="{before}"]', pane)
    dest = waitElement(By.XPATH, f'//ListItem[contains(@Name, "{before}")]', pane)
    ppt.scroll(dest, dest)
    pane.send_keys(Keys.ARROW_RIGHT)
    pane.send_keys(Keys.DELETE)
    pane.send_keys(Keys.ARROW_LEFT)
    pane.send_keys(Keys.CONTROL + 'v')
    pane.send_keys(Keys.CONTROL)
    pane.send_keys('k')


def saveNewSlide(ppt, pane, output):
    pane.send_keys(Keys.F12)
    sw = ppt.find_element_by_xpath('//Window[@Name="????????????"]')
    sw.send_keys(output)
    sw.find_element_by_xpath('//Button[@Name="??????(S)"]').click()

    try:
        cb = sw.find_elements_by_xpath('//Window[@Name="??????????????????"]//Button[@Name="???(Y)"]')
        if 0 < len(cb):
            cb[0].click()
    except (NoSuchElementException, StaleElementReferenceException):
        print('no need to confirm override')


def prepareFolder():
    f = d1.find_element_by_name('??????????????????')
    f.click()
    f.send_keys(Keys.ENTER)

    fWin = waitElement(By.XPATH, '//Window[@ClassName="CabinetWClass" and @Name="??????????????????"]')
    folder = getDriverFromWin(fWin)

    ps = folder.find_elements_by_xpath('//ListItem')

    mp = partial(matchP, ps)
    pVals = list(map(mp, pKeys))
    return folder, pVals


def findPptx(existingChrome=False):
    if not existingChrome:
        c = d1.find_element_by_name('???????????? - Chrome')
        c.click()
        c.send_keys(Keys.ENTER)

    cWin = waitElement(By.XPATH, '//Pane[@ClassName="Chrome_WidgetWin_1"]')
    chrome = getDriverFromWin(cWin)
    # dumpPageSource(chrome, "chrome-inspect.xml")
    doc = None
    if not existingChrome:
        chrome.find_element_by_xpath('//Button[@Name="????????????????????? - Google ????????????"]').click()
        doc = waitElement(By.XPATH, '//Document[@Name="????????????????????? - Google ????????????"]', chrome)
        try:
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%m%d'), chrome)
        except NoSuchElementException:
            f = None
        if f:
            f.click()
            f.send_keys(Keys.ENTER)
        else:
            f = doc.find_element_by_xpath('//DataItem//Text[@Name = "????????????????????????"]')
            f.click()
            f.send_keys(Keys.ENTER)
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%Y'), doc)
            f.click()
            f.send_keys(Keys.ENTER)
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%m%d'), doc)
            f.click()
            f.send_keys(Keys.ENTER)
    return chrome, doc, existingChrome


def downloadPptx(chrome, doc, existingChrome):
    if not existingChrome:
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pKeys[0], doc)
        p.click()
        p.send_keys(Keys.SHIFT + Keys.F10)
        waitElement(By.XPATH, '//MenuItem[contains(@Name, "??????")]', doc).click()
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pKeys[1], doc)
        p.click()
        p.send_keys(Keys.SHIFT + Keys.F10)
        waitElement(By.XPATH, '//MenuItem[contains(@Name, "??????")]', doc).click()
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pKeys[2], doc)
        p.click()
        p.send_keys(Keys.SHIFT + Keys.F10)
        waitElement(By.XPATH, '//MenuItem[contains(@Name, "??????")]', doc).click()
    # waitElement(By.XPATH, '/Pane/Pane/Pane/Button[contains(@Name, ".ppt")]', chrome)
    waitElement(By.XPATH, '//Pane[@Name="Google Chrome"]//Button[contains(@Name, ".ppt")]', chrome)
    # waitElement(By.XPATH, '//Document[@Name="Downloads"]//DataItem[@AutomationId="title-area"]', chrome)
    # dl = chrome.find_elements_by_xpath('//Document[@Name="Downloads"]//DataItem[@AutomationId="title-area"]')
    dl = chrome.find_elements_by_xpath('//Pane[@Name="Google Chrome"]//Button[contains(@Name, ".ppt")]')
    mp = partial(matchP, dl)
    pVals = list(map(mp, pKeys))
    return chrome, pVals, existingChrome


def getNumOfSlides(ppt):
    ind = ppt.find_element_by_xpath('//Button[contains(@Name, "???????????? ?????????")]')
    return int(ind.get_attribute('Name').replace('???????????? ????????? 1 / ', ''))


def mergePptx(container, pvals, existingChrome):
    pNames = list(map(lambda x: x.get_attribute('Name'), pvals))

    subject = re.sub(r'^(.+)\.([^\s]+)( \([0-9]+\))?\.pptx?$', r'????????? \2', pNames[2]).strip()
    output = re.sub(r'^(.+)_(\w+)\.google?????????( \([0-9]+\))?\.pptx$', r'\1_????????????', pNames[1]).strip()

    ppt3 = launchPpt(pKeys[2], pvals[2], container)
    p3 = toggleSlideView(ppt3)
    copyAllSlides(ppt3, p3)
    closeWindow(ppt3)
    switchWindow(container)

    ppt2 = launchPpt(pKeys[1], pvals[1], container)
    p2 = toggleSlideView(ppt2)
    insertSlides(ppt2, p2, subject)
    copyAllSlides(ppt2, p2)
    closeWindow(ppt2, True)
    switchWindow(container)

    ppt1 = launchPpt(pKeys[0], pvals[0], container)
    numSlides = getNumOfSlides(ppt1)
    p1 = toggleSlideView(ppt1)
    appendSlides(ppt1, p1)
    customPresentation(ppt1, numSlides)
    saveNewSlide(ppt1, p1, output)
    closeWindow(ppt1)
    switchWindow(container)


def customPresentation(ppt, after=0, before=None):
    ppt.find_element_by_xpath('//TabItem[@Name="???????????????"]').click()
    ppt.find_element_by_xpath('//MenuItem[@Name="?????????????????????"]').click()
    ppt.find_element_by_xpath('//MenuItem[@Name="????????????..."]').click()
    ppt.find_element_by_xpath('//Window[@Name="????????????"]//Button[@Name="??????..."]').click()
    ppt.find_element_by_xpath('//Window[@Name="??????????????????"]//Edit[@Name="?????????????????????"]').send_keys("????????????")
    slides = ppt.find_elements_by_xpath('//Window[@Name="??????????????????"]//ListItem')
    for slide in slides[after:before]:
        slide.click()
        slide.send_keys(Keys.SPACE)
    ppt.find_element_by_xpath('//Window[@Name="??????????????????"]//Button[@Name="??????"]').click()
    ppt.find_element_by_xpath('//Window[@Name="??????????????????"]//Button[@Name="??????"]').click()
    ppt.find_element_by_xpath('//Window[@Name="????????????"]//Button[@Name="??????"]').click()


def uploadMergedPptx(chrome, doc, existingChrome):
    doc = waitElement(By.XPATH, '//Document[contains(@Name, "Google ????????????")]', chrome)
    doc.find_element_by_xpath('//MenuItem[@Name="??????"]').click()
    waitElement(By.XPATH, '//MenuItem[@Name="????????????"]', doc).click()
    owin = waitElement(By.XPATH, '//Window[@Name="??????"]', chrome)
    owin.find_element_by_xpath('//TreeItem[@Name="?????? (?????????)"]').click()
    file = owin.find_element_by_xpath('//ListItem[contains(@Name, "????????????")]')
    file.click()
    file.send_keys(Keys.ENTER)
    try:
        waitElement(By.XPATH, '//Button[@Name = "??????"]', doc).click()
    except NoSuchElementException:
        pass
    except TimeoutException:
        pass
    waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % '????????????', doc)


def publishDataSheet(chrome, doc, existingChrome):
    if not existingChrome:
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % '?????????', doc)
        p.click()
        p.send_keys(Keys.ENTER)
    doc = waitElement(By.XPATH, '//Document[contains(@Name, "?????????")]', chrome)
    # waitElement(By.XPATH, '//MenuItem[@Name="????????????"]', doc)
    # waitElement(By.XPATH, '//MenuItem[@Name="??????"]', doc).click()
    # try:
    #     waitElement(By.XPATH, '//MenuItem[@Name="??????????????? w"]', doc).click()
    # except (TimeoutException, NoSuchElementException):
    #     doc.send_keys(Keys.ALT+'f')
    #     doc.send_keys('w')
    # d = waitElement(By.XPATH, '//Custom[@Name="???????????????"]', doc)
    # try:
    #     d.find_element_by_xpath('//Button[@Name="??????"]').click()
    #     waitElement(By.XPATH, '//Custom[@Name="Google Drive"]//Button[@Name="OK"]', chrome).click()
    # except NoSuchElementException:
    #     print('Published already')
    # d.find_element_by_xpath('//Button[@Name="??????"]').click()
    urlBar = chrome.find_element_by_xpath('//Edit[@Name="??????????????????"]')
    url = urlBar.get_attribute('Value.Value')
    return re.sub(r'(https://)?docs\.google\.com/spreadsheets/d/(.+)/edit(#gid=[0-9]+)?', r'\2', url)


def extractSubject(chrome, doc, existingChrome):
    # look for local config
    with open(weeklyConfig.replace('config.json', 'weekly.json'), 'r', encoding='utf-8') as f:
        config = json.load(f)
        dt = [int(x) for x in config['??????'].split('/')]
        sun = getSunday()
        if existingChrome or dt[0] == sun.month and dt[1] == sun.day:
            return config['??????'], config['??????']
    # use final file name on google drive if config not present
    p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pKeys[2], doc)
    return re.sub(r'^(.+)\.(\w+)( \([0-9]+\))?\.pptx$', r'\2', p.get_attribute('Name')).strip()


def getSunday():
    today = datetime.date.today()
    dow = (today.weekday() + 1) % 7
    return today + datetime.timedelta((7 - dow) % 7)


def getThursday():
    today = datetime.date.today()
    dow = (today.weekday() + 1) % 7
    return today + datetime.timedelta((4 - dow) % 7)


def writeWeeklyConfig(sheetId):
    with open(weeklyConfig, 'r+', encoding='utf-8') as f:
        config = json.load(f)
        f.seek(0)
        config['models'][0]['key'] = f'{sheetId}'
        json.dump(config, f, indent=2)
        f.truncate()
        f.flush()


def youtubeSetup(subject, preach, chrome, doc, existingChrome):
    thumb = downloadYoutubeThumbnail(chrome, doc, existingChrome)
    scheduleYoutube(subject, preach, '??????????????????', getSunday(), '??????10:30', chrome, doc, existingChrome, thumb)
    # if not existingChrome:  # swich to a different page for a clean cleanstart
    #     chrome.find_element_by_xpath('//Button[@Name="????????????"]').click()
    # scheduleYoutube(subject, preach, '?????????', getThursday(), '??????8:00', chrome, doc, existingChrome)

def downloadYoutubeThumbnail(chrome, doc, existingChrome):
    if not existingChrome:
        chrome.find_element_by_xpath('//Button[@Name="????????????????????? - Google ????????????"]').click()
        doc = waitElement(By.XPATH, '//Document[@Name="????????????????????? - Google ????????????"]', chrome)
        try:
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%m%d'), chrome)
        except NoSuchElementException:
            f = None
        if f:
            f.click()
            f.send_keys(Keys.ENTER)
        else:
            f = doc.find_element_by_xpath('//DataItem//Text[@Name = "????????????????????????"]')
            f.click()
            f.send_keys(Keys.ENTER)
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%Y'), doc)
            f.click()
            f.send_keys(Keys.ENTER)
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%m%d'), doc)
            f.click()
            f.send_keys(Keys.ENTER)

        try:
            p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % '????????????', doc)
            p.click()
            p.send_keys(Keys.SHIFT + Keys.F10)
            waitElement(By.XPATH, '//MenuItem[contains(@Name, "??????")]', doc).click()
            return True
        except NoSuchElementException:
            return False
        except TimeoutException:
            return False

def scheduleYoutube(subject, preach, key, date, time, chrome, doc, existingChrome, thumb=False):
    if not existingChrome:
        chrome.find_element_by_xpath('//Button[@Name="?????? - YouTube Studio"]').click()
    doc = waitElement(By.XPATH, '//Document[@Name="?????? - YouTube Studio"]', chrome)
    waitElement(By.XPATH, '//Button[@Name="??????????????????"]', doc).click()
    waitElement(By.XPATH, '//ListItem[contains(@Name,"Video thumbnail.")]', doc).click()
    waitElement(By.XPATH, '//ListItem[contains(@Name,"Video thumbnail.") and contains(@Name, "%s")]' % key, doc).click()
    doc.find_element_by_xpath('//Button[@Name="REUSE SETTINGS"]').click()
    # doc.find_element_by_xpath('//Button[@Name="????????????"]').click()
    if key == '??????????????????':
        t = waitElement(By.XPATH, '//Group[@Name="????????????????????????????????????"]', doc)
        ytSubject = '??????????????????%s???%s???%s??????' % (date.strftime('%Y.%m.%d'), subject, preach)
        t.click()
        t.send_keys(Keys.CONTROL + 'a')
        t.send_keys(ytSubject)
        if thumb:
            doc.send_keys(Keys.TAB)
            doc.send_keys(Keys.TAB)
            doc.send_keys(Keys.TAB)
            doc.send_keys(Keys.TAB)
            doc.send_keys(Keys.TAB)
            waitElement(By.XPATH, '//Button[@Name="??????"]', doc).click()
            waitElement(By.XPATH, '//MenuItem[@Name="??????"]', doc).click()

            owin = waitElement(By.XPATH, '//Window[@Name="??????"]', chrome)
            owin.find_element_by_xpath('//TreeItem[@Name="?????? (?????????)"]').click()
            file = owin.find_element_by_xpath('//ListItem[contains(@Name, "????????????")]')
            file.click()
            file.send_keys(Keys.ENTER)

    waitElement(By.XPATH, '//Button[@Name="??????"]', doc).click()
    waitElement(By.XPATH, '//Button[@Name="??????"]', doc).click()
    oneMoreTab = False
    try:
        waitElement(By.XPATH, '//RadioButton[@Name="?????????"]', doc).click()
        doc.send_keys(Keys.TAB)
        doc.send_keys(Keys.ENTER)
        # waitElement(By.XPATH, '//*[@AutomationId="datepicker-trigger"]', doc).click()
        d = waitElement(By.XPATH, '//Group[@Name="????????????"]//Edit', doc)
    except TimeoutException:
        oneMoreTab = True
        waitElement(By.XPATH, '//RadioButton[@Name="?????????"]', doc).click()
        doc.send_keys(Keys.TAB)
        doc.send_keys(Keys.TAB)
        doc.send_keys(Keys.ENTER)
        d = waitElement(By.XPATH, '//Group[@Name="????????????"]//Edit', doc)
    d.send_keys(Keys.CONTROL + 'a')
    d.send_keys(date.strftime('%Y???%m???%d???'))
    d.send_keys(Keys.ENTER)

    waitElement(By.XPATH, '//RadioButton[@Name="?????????"]', doc).click()
    if oneMoreTab:
        doc.send_keys(Keys.TAB)
    doc.send_keys(Keys.TAB)
    doc.send_keys(Keys.TAB)
    # doc.send_keys(Keys.ENTER)
    # waitElement(By.XPATH, '//*[@AutomationId="time-of-day-trigger"]', doc).click()
    # waitElement(By.XPATH, '//ListItem[@Name="%s"]' % time, doc).click()
    waitElement(By.XPATH, '//Group[@AutomationId="textbox"]', doc).click()
    waitElement(By.XPATH, '//ListItem[@Name="%s"]' % time, doc).click()

    doc.find_element_by_xpath('//Button[@Name="??????"]').click()
    if key == '??????????????????':
        doc = waitElement(By.XPATH, '//Document[@Name="?????? - YouTube Studio"]', chrome)
        chat = waitElement(By.XPATH, '//Group[@Name="????????????..."]', doc)
        chat.click()
        chat.send_keys('''???????????? ?????????????????????????????????~
        ??????????????????: https://ngpc.tw/weekly/
        ??????????????????: https://ngpc.tw/forms/dedication.html''')
        chat.send_keys(Keys.ENTER)
        waitElement(By.XPATH, f'//Button[@Name="???????????????"]', doc).send_keys(Keys.SPACE)
        # waitElement(By.XPATH, f'//Button[@Name="????????????"]', doc).send_keys(Keys.SPACE)
        waitElement(By.XPATH, '//ListItem[@Name="???????????????"]', doc).click()


def legacyYoutubeSetup(subject, chrome, doc, existingChrome):
    ytSubject = '??????????????????%s???%s??????????????????' % (getSunday().strftime('%Y.%m.%d'), subject)

    if not existingChrome:
        chrome.find_element_by_xpath('//Button[@Name="?????? - YouTube Studio"]').click()
    doc = waitElement(By.XPATH, '//Document[@Name="?????? - YouTube Studio"]', chrome)
    waitElement(By.XPATH, '//Button[@Name="??????????????????"]', doc)
    waitElement(By.XPATH, '//Button[@Name="??????"]', doc).click()
    dialog = waitElement(By.NAME, '????????????', doc)
    t = dialog.find_element_by_xpath('//Group[@AutomationId="textbox" and @Name="????????????????????????????????????"]')
    if t.get_attribute('Value.Value') != ytSubject:
        t.click()
        t.send_keys(Keys.CONTROL + 'a')
        t.send_keys(ytSubject)
        doc.find_element_by_xpath('//Button[@Name="??????"]').click()
    else:
        # Cancel if no change made
        doc.find_element_by_xpath('//Button[@Name="??????"]').click()

    existing = 1
    try:
        # Remove existing pin if exists
        dots = doc.find_element_by_xpath('//Button[@Name="????????????"]')
        dots.send_keys(Keys.SPACE)
        waitElement(By.XPATH, '//ListItem[@Name="??????????????????"]', doc).click()
        dots.send_keys(Keys.SPACE)
        waitElement(By.XPATH, '//ListItem[@Name="??????"]', doc).click()
        existing += 1
    except NoSuchElementException:
        print('No existing pin')

    chat = doc.find_element_by_xpath('//Group[@Name="????????????..."]')
    chat.click()
    chat.send_keys('''???????????? ?????????????????????????????????~
    ??????????????????: https://ngpc.tw/weekly/
    ??????????????????: https://ngpc.tw/forms/dedication.html''')
    chat.send_keys(Keys.ENTER)
    waitElement(By.XPATH, f'//Button[@Name="????????????"][{existing}]', doc).send_keys(Keys.SPACE)
    waitElement(By.XPATH, '//ListItem[@Name="???????????????"]', doc).click()


def syncHymnsDb(existingChrome=False):
    if not existingChrome:
        c = d1.find_element_by_name('???????????? - Chrome')
        c.click()
        c.send_keys(Keys.ENTER)

    cWin = waitElement(By.XPATH, '//Pane[@ClassName="Chrome_WidgetWin_1"]')
    chrome = getDriverFromWin(cWin)
    if not existingChrome:
        chrome.find_element_by_xpath('//Button[@Name="????????????????????? - Google ????????????"]').click()
        doc = waitElement(By.XPATH, '//Document[@Name="????????????????????? - Google ????????????"]', chrome)
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % '???????????????', doc)
        p.click()
        p.send_keys(Keys.ENTER)
    doc = waitElement(By.XPATH, '//Document[contains(@Name, "???????????????")]', chrome)
    waitElement(By.XPATH, '//MenuItem[@Name="????????????"]', doc).click()
    waitElement(By.XPATH, '//MenuItem[@Name="???????????????????????????"]', doc).click()
    waitElement(By.XPATH, '//Button[@Name="??????"]', doc).click()
    return chrome, doc, existingChrome


def setupWindows(screen, existingChrome=False):
    print(screen)
    if not existingChrome:
        c = d1.find_element_by_name('???????????? - Chrome')
        c.click()
        c.send_keys(Keys.ENTER)

    cWin = waitElement(By.XPATH, '//Pane[@ClassName="Chrome_WidgetWin_1"]')
    print(cWin.size)
    chrome = getDriverFromWin(cWin)
    # chrome.find_element_by_xpath('//Button[@Name="????????????????????? - Google ????????????"]').click()
    # doc = waitElement(By.XPATH, '//Document[@Name="????????????????????? - Google ????????????"]', chrome)
    cWin.set_window_size(round(screen['width'] / 3), round(screen['height'] / 2))
    cWin.set_window_position(round(screen['width'] * 2 / 3), 0)


def sendNotificationEmail(chrome, email, subject, body, ccadmin=False):
    if chrome is None:
        print('No browser context found, unable to send email.')
        return
    if email is None or email == '':
        print('No email recipient specified.')
        return
    subject = urllib.parse.quote_plus(subject)
    body = urllib.parse.quote_plus(body)
    urlBar = chrome.find_element_by_xpath('//Edit[@Name="??????????????????"]')
    cc = f'&cc={os.environ["ADMIN"]}' if ccadmin and 'ADMIN' in os.environ else ''
    urlBar.click()
    urlBar.send_keys(f'https://mail.google.com/mail/u/0/?view=cm&fs=1&tf=1&to={email}&su={subject}&body={body}{cc}')
    # urlBar.send_keys(f'mailto:{email}?subject={subject}&body={body}{cc}')
    urlBar.send_keys(Keys.ENTER)
    waitElement(By.XPATH, '//Document[@Name="???????????? - ngpc0706@gmail.com - Gmail"]', chrome).send_keys(
        Keys.CONTROL + Keys.ENTER)
    time.sleep(5)


def getUrl(chrome, doc, existingChrome):
    urlBar = chrome.find_element_by_xpath('//Edit[@Name="??????????????????"]')
    return urlBar.get_attribute('Value.Value')


def dumpPageSource(driver, file='dump.xml'):
    with open(file, 'w', encoding='utf-8') as f:
        f.write(driver.page_source)


def closeWindow(win, saveNot=False):
    win.close_app()
    # win.find_element_by_xpath('//Button[@Name="??????"]').click()
    if saveNot:
        nsBtn = win.find_elements_by_xpath('//Button[@Name="????????????"]')
        if 0 < len(nsBtn):
            nsBtn[0].click()
    win.quit()


taskNames = {
    'weeklyPub': '??????????????????',
    'mergePptx': '?????????????????????',
    'youtubeSetup': '??????Youtube????????????',
    'hymnsDbSync': '?????????????????????',
}
pKeys = ['??????', '????????????.google?????????', '??????']
taskChoices = taskNames.keys()
weeklyConfig = '../ngpc/models/config.json'
parser = argparse.ArgumentParser(description='NGPC church automation tool.')
parser.add_argument('--task', action='store', default='', required=True, choices=taskChoices,
                    help='Choose either one of the tasks')
parser.add_argument('--email', action='store', default='',
                    help='Email to send notification to')
parser.add_argument('--weekly-config', action='store', dest=weeklyConfig,
                    help='Config path of weeklyPub to write to')
parser.add_argument('--dry-run', action='store_true',
                    help='Specify if don\'t want to commit result yet')

args = parser.parse_args()

if args.task in taskChoices:
    print(f'Running {args.task}')
    # set up appium
    root = launchRoot()
    d1 = root.find_element_by_name('?????? 1')
    d1.send_keys(Keys.COMMAND + 'd')
    context = None
    try:
        resultUrl = ''
        if 'weeklyPub' == args.task:
            context = findPptx()
            sid = publishDataSheet(*context)
            writeWeeklyConfig(sid)
            subprocess.run(['npm', 'run', 'prepare'], check=True, shell=True, cwd='../ngpc')
            if not args.dry_run:
                subprocess.run(['npm', 'run', 'upload'], check=True, shell=True, cwd='../ngpc')
            resultUrl = 'https://ngpc.tw/weekly/'
        elif 'mergePptx' == args.task:
            context = findPptx()
            context2 = downloadPptx(*context)
            mergePptx(*context2)
            if not args.dry_run:
                uploadMergedPptx(*context)
            resultUrl = getUrl(*context)
        elif 'youtubeSetup' == args.task:
            context = findPptx()
            sid = publishDataSheet(*context)
            writeWeeklyConfig(sid)
            subprocess.run(['npm', 'run', 'loadWeekly'], check=True, shell=True, cwd='../ngpc')
            subj, prch = extractSubject(*context)
            if not args.dry_run:
                youtubeSetup(subj, prch, *context)
            resultUrl = getUrl(*context)
        elif 'hymnsDbSync' == args.task:
            context = syncHymnsDb()
            resultUrl = getUrl(*context)
        # elif 'liveSetup' == args.task:
        #     setupWindows(d1.size)
        sendNotificationEmail(
            context[0], args.email,
            f'?????????????????? {taskNames[args.task]} ??????????????????',
            f'''{taskNames[args.task]} ?????????????????????\n\n????????????????????????\n{resultUrl}''')
    except (TimeoutException, NoSuchElementException) as e:
        sendNotificationEmail(
            context[0], args.email,
            f'???????????????????????? {taskNames[args.task]} ??????????????????',
            f'{taskNames[args.task]} ?????????????????????????????????\n???????????????????????????????????????',
            True)
        raise e
    except subprocess.CalledProcessError as e:
        sendNotificationEmail(
            context[0], args.email,
            f'???????????????????????? {taskNames[args.task]} ??????????????????',
            f'{taskNames[args.task]} ?????????????????????????????????????????????????????????\n???????????????????????????????????????????????????????????????',
            True)
        raise e
    finally:
        closeWindow(context[0])
        print(f'Finished task: {args.task}')
        root.quit()
