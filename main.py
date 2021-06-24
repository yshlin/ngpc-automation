from selenium.common.exceptions import NoSuchElementException
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
    ebtn = ppt.find_elements_by_xpath('//Button[@Name="啟用編輯(E)"]')
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
    ppt.find_element_by_xpath('//TabItem[@Name="檢視"]').click()
    ppt.find_element_by_xpath('//Button[@Name="投影片瀏覽"]').click()
    return ppt.find_element_by_xpath('//Pane[@Name="投影片瀏覽"]')


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
    dest = ppt.find_element_by_xpath(f'//Pane[@Name="投影片瀏覽"]//ListItem[@Name="{before}"]')
    ppt.scroll(dest, dest)
    pane.send_keys(Keys.ARROW_RIGHT)
    pane.send_keys(Keys.DELETE)
    pane.send_keys(Keys.ARROW_LEFT)
    pane.send_keys(Keys.CONTROL + 'v')
    pane.send_keys(Keys.CONTROL)
    pane.send_keys('k')


def dumpPageSource(driver, file='dump.xml'):
    with open(file, 'w', encoding='utf-8') as f:
        f.write(driver.page_source)


def closeWindow(win, saveNot=False):
    win.close_app()
    # win.find_element_by_xpath('//Button[@Name="關閉"]').click()
    if saveNot:
        nsbtn = win.find_elements_by_xpath('//Button[@Name="不要儲存"]')
        if 0 < len(nsbtn):
            nsbtn[0].click()
    win.quit()


def saveNewSlide(ppt, pane, output):
    pane.send_keys(Keys.F12)
    sw = ppt.find_element_by_xpath('//Window[@Name="另存新檔"]')
    sw.send_keys(output)
    sw.find_element_by_xpath('//Button[@Name="儲存(S)"]').click()

    try:
        cb = sw.find_elements_by_xpath('//Window[@Name="確認另存新檔"]//Button[@Name="是(Y)"]')
        if 0 < len(cb):
            cb[0].click()
    except NoSuchElementException:
        pass  # no need to confirm override


def prepareFolder():
    f = d1.find_element_by_name('當週主日投影')
    f.click()
    f.send_keys(Keys.ENTER)

    fWin = waitElement(By.XPATH, '//Window[@ClassName="CabinetWClass" and @Name="當週主日投影"]')
    folder = getDriverFromWin(fWin)

    ps = folder.find_elements_by_xpath('//ListItem')

    mp = partial(matchP, ps)
    pVals = list(map(mp, pkeys))
    return folder, pVals


def mergePptx(container, pvals, existingChrome):
    pNames = list(map(lambda x: x.get_attribute('Name'), pvals))

    subject = re.sub(r'^(.+)\.(\w+)( \([0-9]+\))?\.pptx$', r'投影片 \2', pNames[2]).strip()
    output = re.sub(r'^(.+)_(\w+)\.google簡報檔( \([0-9]+\))?\.pptx$', r'\1_自動合併', pNames[1]).strip()

    ppt3 = launchPpt(pkeys[2], pvals[2], container)
    p3 = toggleSlideView(ppt3)
    copyAllSlides(ppt3, p3)
    closeWindow(ppt3)
    switchWindow(container)

    ppt2 = launchPpt(pkeys[1], pvals[1], container)
    p2 = toggleSlideView(ppt2)
    insertSlides(ppt2, p2, subject)
    copyAllSlides(ppt2, p2)
    closeWindow(ppt2, True)
    switchWindow(container)

    ppt1 = launchPpt(pkeys[0], pvals[0], container)
    p1 = toggleSlideView(ppt1)
    appendSlides(ppt1, p1)
    customPresentation(ppt1, 10)
    saveNewSlide(ppt1, p1, output)
    closeWindow(ppt1)
    switchWindow(container)

    # closeWindow(container)


def findPptx(existingChrome=False):
    if not existingChrome:
        c = d1.find_element_by_name('南園教會 (NGPC) - Chrome')
        c.click()
        c.send_keys(Keys.ENTER)

    cWin = waitElement(By.XPATH, '//Pane[@ClassName="Chrome_WidgetWin_1"]')
    chrome = getDriverFromWin(cWin)
    # dumpPageSource(chrome, "chrome-inspect.xml")
    doc = None
    if not existingChrome:
        chrome.find_element_by_xpath('//Button[@Name="影音自動化樣板 - Google 雲端硬碟"]').click()
        doc = waitElement(By.XPATH, '//Document[@Name="影音自動化樣板 - Google 雲端硬碟"]', chrome)
        try:
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%m%d'), chrome)
        except NoSuchElementException:
            f = None
        if f:
            f.click()
            f.send_keys(Keys.ENTER)
        else:
            f = doc.find_element_by_xpath('//DataItem//Text[@Name = "過去影音文件備份"]')
            f.click()
            f.send_keys(Keys.ENTER)
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%Y'), doc)
            f.click()
            f.send_keys(Keys.ENTER)
            f = waitElement(By.XPATH, '//DataItem//Text[@Name = "%s"]' % getSunday().strftime('%m%d'), doc)
            f.click()
            f.send_keys(Keys.ENTER)
    return chrome, doc, existingChrome


def extractSubject(chrome, doc, existingChrome):
    p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pkeys[2], doc)
    return re.sub(r'^(.+)\.(\w+)( \([0-9]+\))?\.pptx$', r'\2', p.get_attribute('Name')).strip()


def downloadPptx(chrome, doc, existingChrome):
    if not existingChrome:
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pkeys[0], doc)
        p.click()
        p.send_keys(Keys.SHIFT + Keys.F10)
        waitElement(By.XPATH, '//MenuItem[contains(@Name, "下載")]', doc).click()
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pkeys[1], doc)
        p.click()
        p.send_keys(Keys.SHIFT + Keys.F10)
        waitElement(By.XPATH, '//MenuItem[contains(@Name, "下載")]', doc).click()
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % pkeys[2], doc)
        p.click()
        p.send_keys(Keys.SHIFT + Keys.F10)
        waitElement(By.XPATH, '//MenuItem[contains(@Name, "下載")]', doc).click()
    waitElement(By.XPATH, '/Pane/Pane/Pane/Button[contains(@Name, ".pptx")]', chrome)
    # waitElement(By.XPATH, '//Document[@Name="Downloads"]//DataItem[@AutomationId="title-area"]', chrome)
    # dl = chrome.find_elements_by_xpath('//Document[@Name="Downloads"]//DataItem[@AutomationId="title-area"]')
    dl = chrome.find_elements_by_xpath('/Pane/Pane/Pane/Button[contains(@Name, ".pptx")]')
    mp = partial(matchP, dl)
    pVals = list(map(mp, pkeys))
    return chrome, pVals, existingChrome


def youtubeSetup(subject, chrome, doc, existingChrome):
    ytSubject = '【週日禮拜】%s《%s》李俊佑牧師' % (getSunday().strftime('%Y.%m.%d'), subject)

    chrome.find_element_by_xpath('//Button[@Name="直播 - YouTube Studio"]').click()
    doc = waitElement(By.XPATH, '//Document[@Name="直播 - YouTube Studio"]', chrome)
    clink = doc.find_elements_by_xpath('//Hyperlink[@Name="關閉"]')
    if 0 < len(clink):
        clink[0].click()
    doc.find_element_by_xpath('//Button[@Name="編輯"]').click()
    dialog = waitElement(By.NAME, '編輯設定', chrome)
    t = dialog.find_element_by_xpath('//Group[@AutomationId="textbox" and @Name="新增可描述直播內容的標題"]')
    t.send_keys(ytSubject)
    closeWindow(chrome)


def customPresentation(ppt, after=0, before=None):
    ppt.find_element_by_xpath('//TabItem[@Name="投影片放映"]').click()
    ppt.find_element_by_xpath('//MenuItem[@Name="自訂投影片放映"]').click()
    ppt.find_element_by_xpath('//MenuItem[@Name="自訂放映..."]').click()
    ppt.find_element_by_xpath('//Window[@Name="自訂放映"]//Button[@Name="新增..."]').click()
    ppt.find_element_by_xpath('//Window[@Name="定義自訂放映"]//Edit[@Name="投影片放映名稱"]').send_keys("主日禮拜")
    slides = ppt.find_elements_by_xpath('//Window[@Name="定義自訂放映"]//ListItem')
    for slide in slides[10:]:
        slide.click()
        slide.send_keys(Keys.SPACE)
    ppt.find_element_by_xpath('//Window[@Name="定義自訂放映"]//Button[@Name="新增"]').click()
    ppt.find_element_by_xpath('//Window[@Name="定義自訂放映"]//Button[@Name="確定"]').click()
    ppt.find_element_by_xpath('//Window[@Name="自訂放映"]//Button[@Name="關閉"]').click()


def uploadMergedPptx(chrome, doc, existingChrome):
    doc = waitElement(By.XPATH, '//Document[contains(@Name, "Google 雲端硬碟")]', chrome)
    doc.find_element_by_xpath('//MenuItem[@Name="新增"]').click()
    waitElement(By.XPATH, '//MenuItem[@Name="檔案上傳"]', doc).click()
    owin = waitElement(By.XPATH, '//Window[@Name="開啟"]', chrome)
    owin.find_element_by_xpath('//TreeItem[@Name="下載 (已釘選)"]').click()
    file = owin.find_element_by_xpath('//ListItem[contains(@Name, "自動合併")]')
    file.click()
    file.send_keys(Keys.ENTER)
    waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % '自動合併', doc)
    closeWindow(chrome)


def publishDataSheet(chrome, doc, existingChrome):
    if not existingChrome:
        p = waitElement(By.XPATH, '//DataItem//Text[contains(@Name, "%s")]' % '產生器', doc)
        p.click()
        p.send_keys(Keys.ENTER)
    doc = waitElement(By.XPATH, '//Document[contains(@Name, "產生器")]', chrome)
    doc.find_element_by_xpath('//MenuItem[@Name="檔案"]').click()
    waitElement(By.XPATH, '//MenuItem[@Name="發布到網路 w"]', doc).click()
    d = waitElement(By.XPATH, '//Custom[@Name="發布到網路"]', doc)
    try:
        d.find_element_by_xpath('//Button[@Name="發布"]').click()
        waitElement(By.XPATH, '//Custom[@Name="Google Drive"]//Button[@Name="OK"]', chrome).click()
    except NoSuchElementException:
        print('Published already')
    d.find_element_by_xpath('//Button[@Name="關閉"]').click()
    urlBar = chrome.find_element_by_xpath('//Edit[@Name="Address and search bar"]')
    url = urlBar.get_attribute('Value.Value')
    closeWindow(chrome)
    return re.sub(r'(https://)?docs\.google\.com/spreadsheets/d/(.+)/edit(#gid=[0-9]+)?', r'\2', url)


def getSunday():
    today = datetime.date.today()
    dow = (today.weekday() + 1) % 7
    return today + datetime.timedelta((7 - dow) % 7)


def writeWeeklyConfig(sheetId):
    with open(weeklyConfig, 'r+') as f:
        config = json.load(f)
        f.seek(0)
        config['models'][0]['key'] = f'{sheetId}/4'
        json.dump(config, f, indent=2)
        f.truncate()


pkeys = ['輪播', '.google簡報檔', '講道']
task = None
taskChoices = ['weeklyPub', 'mergePptx', 'youtubeSetup']
weeklyConfig = '../ngpc/models/config.json'
dryRun = False
parser = argparse.ArgumentParser(description='NGPC church automation tool.')
parser.add_argument('--task', action='store', dest=task, default='', required=True,
                    choices=taskChoices,
                    help='Choose either one of the tasks')
parser.add_argument('--weekly-config', action='store', dest=weeklyConfig,
                    help='Config path of weeklyPub to write to')
parser.add_argument('--dry-run', action='store_true', dest=dryRun,
                    help='Specify if don\'t want to commit result yet')
args = parser.parse_args()

if args.task in taskChoices:
    print(f'Running {args.task}')
    # set up appium
    root = launchRoot()
    d1 = root.find_element_by_name('桌面 1')
    d1.send_keys(Keys.COMMAND + 'd')
    context = findPptx()

    if 'weeklyPub' == args.task:
        sid = publishDataSheet(*context)
        writeWeeklyConfig(sid)
        subprocess.run(['npm', 'run', 'prepare'], check=True, shell=True, cwd='../ngpc')
        if not dryRun:
            subprocess.run(['npm', 'run', 'upload'], check=True, shell=True, cwd='../ngpc')
    elif 'mergePptx' == args.task:
        downloadPptx(*context)
        mergePptx(*context)
        if not dryRun:
            uploadMergedPptx(*context)
    elif 'youtubeSetup' == args.task:
        subj = extractSubject(*context)
        if not dryRun:
            youtubeSetup(subj, *context)

    print(f'Finished task: {args.task}')
    root.quit()
