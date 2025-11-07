import asyncio
import datetime
import os
import schedule
import time
from playwright.async_api import async_playwright
import win32com.client as win32

# --- Credentials ---
ZOHO_EMAIL = "your name.r@xyz.com"
ZOHO_PASSWORD = "your name@28"
ZOHO_URL = "https://accounts.zoho.in/signin?servicename=zohopeople"
OUTLOOK_TO_SUCCESS = "yourname.r@xyz.com"
OUTLOOK_TO_FAIL = "mail name.r@xyz.com"


# --- Outlook Mail Sender ---
def send_outlook_mail(to, subject, body, attachment=None):
    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if attachment and os.path.exists(attachment):
            mail.Attachments.Add(attachment)
        mail.Send()
        print(f"📧 Mail sent to {to}")
    except Exception as e:
        print(f"⚠️ Failed to send mail: {e}")


# --- Zoho Login & Check ---
async def zoho_check(action="checkin"):
    print(f"\n🚀 Starting {action.capitalize()} at {datetime.datetime.now().strftime('%H:%M:%S')}...\n")
    screenshot_path = f"{action}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            context = await browser.new_context(geolocation={"latitude": 12.851492, "longitude": 80.071031},
            permissions=["geolocation"],)
            page = await context.new_page()

            print("🌐 Navigating to Zoho People login...")
            await page.goto(ZOHO_URL, wait_until="networkidle")

            # --- Handle login form ---
            await page.wait_for_timeout(3000)
            frame = None
            for f in page.frames:
                if "zoho" in f.url:
                    frame = f
                    break

            target = frame or page
            print(f"🪟 Using {'iframe' if frame else 'main page'} for login...")

            # --- Login ID ---
            for sel in ['input[name="login_id"]', 'input#login_id', 'input[type="email"]']:
                try:
                    await target.wait_for_selector(sel, timeout=5000)
                    await target.fill(sel, ZOHO_EMAIL)
                    print("✅ Entered email.")
                    break
                except:
                    continue

            #await target.click('button[type="submit"]')
            await page.get_by_role("button", name="Next").click()
            await target.wait_for_timeout(2000)

            # --- Password ---
            for sel in ['input[name="password"]', 'input#password', 'input[type="password"]']:
                try:
                    await target.wait_for_selector(sel, timeout=5000)
                    await target.fill(sel, ZOHO_PASSWORD)
                    print("🔑 Entered password.")
                    break
                except:
                    continue

           #await target.click('button[type="submit"]')
            #await page.get_by_role("button", name="Next").click()
           # sign_in_button = page.locator('button[type="submit"]')
        #await sign_in_button.nth(1).click()
            #await page.wait_for_load_state("networkidle")
            #print("✅ Logged into Zoho successfully.")
            #await page.wait_for_timeout(6000)
        #await target.click('button[type="submit"]')
            #await page.get_by_role("button", name="Next").click()

            #sign_in_button = page.locator('button[type="Signin"]')
            #await sign_in_button.nth(1).click()
            await page.click('xpath=//*[@id="nextbtn"]')
            print("✅ Logged into Zoho successfully.")
            await page.wait_for_timeout(6000)
         # --- Navigate to Attendance ---
            await page.goto("attendance page URL"load")
            await page.wait_for_timeout(5000)

            # --- Perform Check-In/Out ---
            success = False
            
            if action == "checkin":
                    print("🕘 Performing Check-in...")
                    await page.wait_for_selector('xpath=//*[@id="ZPAtt_check_in_out"]/div/p', timeout=30000)
                    button = page.locator('#ZPAtt_check_in_out')
                    await button.click()
                    # await page.click('xpath=//*[@id="ZPAtt_check_in_out"]/div/p')
                    print("🕘 Performed Check-in.")
                    success = True
            else:
                if await page.is_visible("text=Check-out"):
                    await page.click("text=Check-out")
                    print("🕕 Performed Check-out.")
                    await page.wait_for_timeout(50000)
                    success = True
            os.makedirs("screenshots", exist_ok=True)
            screenshot_name = f"{action}_screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            screenshot_path = os.path.join(os.getcwd(), "screenshots", screenshot_name)
            await page.screenshot(path=screenshot_path)
            print(f"📸 Screenshot saved: {screenshot_path}")

            await browser.close()

            # --- Success / Failure Notification ---
    except Exception as e:
        print(f"⚠️ Error during {action}: {e}")
        send_outlook_mail(
            OUTLOOK_TO_FAIL,
            f"Emergency {action.capitalize()} for Mathan",
            f"Hi Selva,\n\nI'm Mathan's AI assistant Beni.\nI tried to {action} Mathan's Zoho today but got this error:\n\n{e}\n\nPlease check manually.\n\nThanks,\nBeni 🤖",
        )

    if success:
        if os.path.exists(screenshot_path):
            try:
                send_outlook_mail(
                    OUTLOOK_TO_SUCCESS,
                    f"✅ Zoho {action.capitalize()} Successful",
                    f"Hi Mathan,\n\nYour Zoho {action} was successful at "
                    f"{datetime.datetime.now().strftime('%H:%M:%S')}.\n\nRegards,\nBeni 🤖",
                    attachment=screenshot_path,
                )
            except Exception as e:
                print(f"⚠️ Failed to send success mail: {e}")
        else:
            print(f"⚠️ Screenshot file not found: {screenshot_path}")
            #     send_outlook_mail(
            #         OUTLOOK_TO_SUCCESS,
            #         f"✅ Zoho {action.capitalize()} Successful",
            #         f"Hi Mathan,\n\nYour Zoho {action} was successful at {datetime.datetime.now().strftime('%H:%M:%S')}.\n\nRegards,\nBeni 🤖",
            #         attachment=screenshot_path,
            #     )
            # else:
            #     raise Exception("Check-in/out button not found or not clickable.")

    # except Exception as e:
    #     print(f"⚠️ Error during {action}: {e}")
    #     send_outlook_mail(
    #         OUTLOOK_TO_FAIL,
    #         f"Emergency {action.capitalize()} for Mathan",
    #         f"Hi Selva,\n\nI’m Mathan’s AI assistant Beni.\nI tried to {action} Mathan’s Zoho today but got this error:\n\n{e}\n\nPlease check manually.\n\nThanks,\nBeni 🤖",
    #     )


# --- Scheduler ---
def schedule_jobs():
    for day in ["monday", "tuesday", "wednesday", "thursday", "friday"]:
        getattr(schedule.every(), day).at("09:31").do(lambda: asyncio.run(zoho_check("checkin")))
        getattr(schedule.every(), day).at("21:55").do(lambda: asyncio.run(zoho_check("checkout")))

    print("🕒 Zoho Scheduler Started\n✔️ Check-in → 09:31 AM\n✔️ Check-out → 06:30 PM\n❌ Skipped on Weekends\n")

    while True:
        schedule.run_pending()
        time.sleep(30)


# --- Run Script ---
if __name__ == "__main__":
    schedule_jobs()
