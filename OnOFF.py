import pathlib
from selenium import webdriver
import chromedriver_autoinstaller
import requests
import os
from dotenv import load_dotenv
from telegram.ext import Updater
from telegram.ext import CommandHandler, Filters, MessageHandler
from telegram import ChatAction
import shutil
import pickle
from os import execl
from sys import executable
import time

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
USER_ID = os.getenv("USER_ID")
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")

scriptDirectory = pathlib.Path().absolute()


options = webdriver.ChromeOptions()
# options.add_argument("--headless")
options.add_argument("--disable-infobars")
options.add_argument("--window-size=1200,800")
# options.add_argument("--disable-web-security")
# options.add_argument("--allow-running-insecure-content")
options.add_argument(f"user-data-dir={scriptDirectory}\\ChromiumData")
options.add_experimental_option(
    "prefs",
    {
        "profile.default_content_setting_values.media_stream_mic": 2,
        "profile.default_content_setting_values.media_stream_camera": 2,
        "profile.default_content_setting_values.notifications": 2,
    },
)

chromedriver_autoinstaller.install()
browser = webdriver.Chrome(options=options)

updater = Updater(token=BOT_TOKEN, use_context=True)
dp = updater.dispatcher


def telegram_bot_sendtext(bot_message):
    send_text = (
        "https://api.telegram.org/bot"
        + BOT_TOKEN
        + "/sendMessage?chat_id="
        + USER_ID
        + "&parse_mode=Markdown&text="
        + bot_message
    )
    requests.get(send_text)


def start(update, context):
    user = update.message.from_user
    # print(user)
    context.bot.send_chat_action(chat_id=user["id"], action=ChatAction.TYPING)
    update.message.reply_text("Hello {}!".format(user["first_name"]))
    update.message.reply_text("Your UserID is: {} ".format(user["id"]))


def status(update, context):
    user = update.message.from_user
    if user["id"] == int(USER_ID):
        browser.save_screenshot("snapshot.png")
        context.bot.send_chat_action(
            chat_id=USER_ID, action=ChatAction.UPLOAD_PHOTO)
        context.bot.send_photo(
            chat_id=USER_ID, photo=open("snapshot.png", "rb"), timeout=100
        )
        os.remove("snapshot.png")
    else:
        update.message.reply_text(
            "You are not authorized to use this bot.\nUse /owner to know about me"
        )


def reset(update, context):
    user = update.message.from_user
    if user["id"] == int(USER_ID):
        context.bot.send_chat_action(
            chat_id=USER_ID, action=ChatAction.TYPING)
        if os.path.exists("ChromiumData") or os.path.exists("outlook.pkl"):

            try:
                browser.quit()
                shutil.rmtree("ChromiumData")
                try:
                    os.remove("outlook.pkl")
                except:
                    pass
                context.bot.send_message(
                    chat_id=USER_ID, text="Chrome Reset Succesfull"
                )
                execl(executable, executable, "OnOFF.py")
            except OSError as e:
                print("Error: %s - %s." % (e.filename, e.strerror))
        else:
            context.bot.send_message(
                chat_id=USER_ID, text="Browser is already clear..."
            )
    else:
        update.message.reply_text(
            "You are not authorized to use this bot.\nUse /owner to know about me"
        )


def help(update, context):
    user = update.message.from_user
    if user["id"] == int(USER_ID):
        context.bot.send_message(
            chat_id=USER_ID,
            text="/online - Set Status Online\n/offline - Set Status offline\n/login - login into teams\n/status - Screenshot of Teams\n/restart - restart the outlookrobot\n/reset - Reset chrome browser\n/owner-To know about me\n/help - To Display this message",
        )
    else:
        update.message.reply_text(
            "You are not authorized to use this bot.\nUse /owner to know about me"
        )


def echo(update, context):
    update.message.reply_text(
        "This is not a valid command\nUse /help to list out the available commands"
    )


def offline(update, context):
    user = update.message.from_user
    if user["id"] == int(USER_ID):
        context.bot.send_message(
            chat_id=USER_ID, text="Going Offline...!")
        pickle.dump("restart msg check", open("restart.pkl", "wb"))
        browser.quit()
        execl(executable, executable, "OnOFF.py")
    else:
        update.message.reply_text(
            "You are not authorized to use this bot.\nUse /owner to know about me"
        )


def online(update, context):
    user = update.message.from_user
    if user["id"] == int(USER_ID):
        context.bot.send_chat_action(chat_id=USER_ID, action=ChatAction.TYPING)
        try:
            if os.path.exists("outlook.pkl"):
                pass
            else:
                context.bot.send_message(
                    chat_id=USER_ID,
                    text="You're not logged in please run /login command to login. Then try again!",
                )
                return

            browser.get("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=id_token&scope=openid%20profile&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=eyJpZCI6IjVkMzNkZWI4LTVlODQtNGYzOS1iODAyLTk5YWY2ODRhMTYyNyIsInRzIjoxNjQzODA3NDQ0LCJtZXRob2QiOiJyZWRpcmVjdEludGVyYWN0aW9uIn0%3D&nonce=3213d044-d2ce-437d-a37c-b22ad6e36a03&client_info=1&x-client-SKU=MSAL.JS&x-client-Ver=1.3.4&client-request-id=d0bd5d43-839a-4f77-8780-536bb267d8d8&response_mode=fragment&sso_reload=true")

            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.TYPING)

            time.sleep(10)

            browser.save_screenshot("screenshot.png")
            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.UPLOAD_PHOTO)
            mid = context.bot.send_photo(
                chat_id=USER_ID, photo=open("screenshot.png", "rb"), timeout=120
            ).message_id
            os.remove("screenshot.png")

            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.TYPING)
            time.sleep(3)
            context.bot.send_message(
                chat_id=USER_ID,
                text="I have set your status to \"Online\" 😉",
            )

        except Exception as e:
            browser.quit()
            context.bot.send_message(
                chat_id=USER_ID, text="Error occurred! Fix error and retry!"
            )
            context.bot.send_message(
                chat_id=USER_ID, text="Try /reset to fix the issue")
            context.bot.send_message(chat_id=USER_ID, text=str(e))
            execl(executable, executable, "OnOFF.py")
    else:
        update.message.reply_text(
            "You are not authorized to use this bot.\nUse /owner to know about me"
        )


def login(update, context):
    user = update.message.from_user
    if user["id"] == int(USER_ID):
        if os.path.exists("outlook.pkl"):
            context.bot.send_message(
                chat_id=USER_ID,
                text="Already Logged In! Run /online to set status \"online\"",
            )
            context.bot.send_message(
                chat_id=USER_ID,
                text="Still getting some error? try using /reset",
            )
            return
        try:
            update.message.reply_text("Logging in...")
            context.bot.send_chat_action(
                chat_id=user["id"], action=ChatAction.TYPING)
            browser.get(
                "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=id_token&scope=openid%20profile&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=eyJpZCI6IjVkMzNkZWI4LTVlODQtNGYzOS1iODAyLTk5YWY2ODRhMTYyNyIsInRzIjoxNjQzODA3NDQ0LCJtZXRob2QiOiJyZWRpcmVjdEludGVyYWN0aW9uIn0%3D&nonce=3213d044-d2ce-437d-a37c-b22ad6e36a03&client_info=1&x-client-SKU=MSAL.JS&x-client-Ver=1.3.4&client-request-id=d0bd5d43-839a-4f77-8780-536bb267d8d8&response_mode=fragment&sso_reload=true"
            )
            time.sleep(3)
            username = browser.find_element_by_id("i0116")
            username.send_keys(OUTLOOK_EMAIL)
            nextButton = browser.find_element_by_id("idSIButton9")
            nextButton.click()
            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.TYPING)
            time.sleep(7)

            browser.save_screenshot("ss.png")
            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.UPLOAD_PHOTO)
            mid = context.bot.send_photo(
                chat_id=USER_ID, photo=open("ss.png", "rb"), timeout=120
            ).message_id
            os.remove("ss.png")

            password = browser.find_element_by_id(
                "i0118")
            password.send_keys(OUTLOOK_PASSWORD)
            signInButton = browser.find_element_by_id("idSIButton9")
            signInButton.click()
            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.TYPING)
            time.sleep(7)

            if browser.find_elements_by_xpath('//*[@id="idSIButton9"]'):
                yesButton = browser.find_element_by_id("idSIButton9")
                yesButton.click()

            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.UPLOAD_PHOTO)
            time.sleep(5)
            browser.save_screenshot("ss.png")
            mid = context.bot.send_photo(
                chat_id=USER_ID, photo=open("ss.png", "rb"), timeout=120
            ).message_id
            os.remove("ss.png")
            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.TYPING)
            time.sleep(7)

            browser.find_element_by_xpath(
                '//*[@id="tilesHolder"]/div[1]/div').click()

            context.bot.send_chat_action(
                chat_id=USER_ID, action=ChatAction.UPLOAD_PHOTO)
            time.sleep(5)
            browser.save_screenshot("ss.png")
            mid = context.bot.send_photo(
                chat_id=USER_ID, photo=open("ss.png", "rb"), timeout=120
            ).message_id
            os.remove("ss.png")
            time.sleep(7)

            context.bot.send_message(
                chat_id=USER_ID,
                text="Logged In Successfully.",
            )

            pickle.dump("Outlook Login: True", open("outlook.pkl", "wb"))

            browser.quit()
            execl(executable, executable, "OnOFF.py")

        except:
            update.message.reply_text("Auto login failed!")
            browser.get(
                "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=id_token&scope=openid%20profile&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=eyJpZCI6IjVkMzNkZWI4LTVlODQtNGYzOS1iODAyLTk5YWY2ODRhMTYyNyIsInRzIjoxNjQzODA3NDQ0LCJtZXRob2QiOiJyZWRpcmVjdEludGVyYWN0aW9uIn0%3D&nonce=3213d044-d2ce-437d-a37c-b22ad6e36a03&client_info=1&x-client-SKU=MSAL.JS&x-client-Ver=1.3.4&client-request-id=d0bd5d43-839a-4f77-8780-536bb267d8d8&response_mode=fragment&sso_reload=true"
            )
            update.message.reply_text("We have opened login page for you")
            update.message.reply_text("Please login manually")
            update.message.reply_text("Don't worry, it's a one time procedure")

    else:
        update.message.reply_text(
            "You are not authorized to use this bot.\nUse /owner to know about me"
        )
    # context.bot.send_chat_action(chat_id=USER_ID, action=ChatAction.TYPING)


def owner(update, context):
    update.message.reply_text(
        "My code lies around the whole Internet 😇\nIt was assembled, modified and upgraded by Pathak Pratik\nSource Code is available here👇\nhttps://github.com/zpratikpathak/",
    )


def main():

    if os.path.exists("restart.pkl"):
        try:
            os.remove("restart.pkl")
            telegram_bot_sendtext("Statuse set to \"Offline\"😁")
        except:
            pass

    dp.add_handler(CommandHandler("start", start, run_async=True))
    dp.add_handler(CommandHandler("help", help, run_async=True))
    dp.add_handler(CommandHandler("offline", offline, run_async=True))
    dp.add_handler(CommandHandler("status", status, run_async=True))
    dp.add_handler(CommandHandler("reset", reset, run_async=True))
    dp.add_handler(CommandHandler("login", login, run_async=True))
    dp.add_handler(CommandHandler("owner", owner, run_async=True))
    dp.add_handler(CommandHandler("online", online, run_async=True))
    dp.add_handler(MessageHandler(Filters.text, echo, run_async=True))

    updater.start_polling()
    updater.idle()


if __name__ == "__main__":
    main()
