import sys
from pathlib import Path
from typing import List

from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException


def wait_for_any_element(driver, locators, timeout=20):
    """Return the first element found from a list of locators."""
    wait = WebDriverWait(driver, timeout)

    def _locate(drv):
        for locator in locators:
            elements = drv.find_elements(*locator)
            if elements:
                return elements[0]
        return False

    return wait.until(_locate)


def wait_for_any_clickable_element(driver, locators, timeout=20):
    """Return the first clickable element found from a list of locators."""
    wait = WebDriverWait(driver, timeout)

    def _locate(drv):
        for locator in locators:
            elements = drv.find_elements(*locator)
            for element in elements:
                if element.is_displayed() and element.is_enabled():
                    return element
        return False

    return wait.until(_locate)


def wait_for_any_elements(driver, locators, timeout=20):
    """Return a non-empty list of elements from the first matching locator."""
    wait = WebDriverWait(driver, timeout)

    def _locate(drv):
        for locator in locators:
            elements = drv.find_elements(*locator)
            if elements:
                return elements
        return False

    return wait.until(_locate)


def try_click_optional(driver, locators, timeout=5) -> bool:
    try:
        element = wait_for_any_clickable_element(driver, locators, timeout=timeout)
        element.click()
        return True
    except TimeoutException:
        return False


def save_instructions_to_excel(instructions: List[str], file_path: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Driving Instructions"
    sheet.append(["Step Number", "Instruction Text"])

    for index, text in enumerate(instructions, start=1):
        sheet.append([index, text])

    workbook.save(file_path)


def save_full_page_screenshot(driver, file_path: str) -> None:
    # Resize to full page dimensions for a true full-page screenshot.
    metrics = driver.execute_cdp_cmd("Page.getLayoutMetrics", {})
    width = int(metrics["contentSize"]["width"])
    height = int(metrics["contentSize"]["height"])
    driver.set_window_size(width, height)
    driver.save_screenshot(file_path)


def main() -> int:
    print("Starting Google Maps route extraction...")

    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")

    driver = None

    try:
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 30)

        print("Opening Google Maps...")
        driver.get("https://maps.google.com")

        try_click_optional(
            driver,
            [
                (By.CSS_SELECTOR, "button[aria-label='Accept all']"),
                (By.CSS_SELECTOR, "button[aria-label='I agree']"),
            ],
            timeout=5,
        )

        directions_button = wait_for_any_clickable_element(
            driver,
            [
                (By.CSS_SELECTOR, "button[data-value='Directions']"),
                (By.CSS_SELECTOR, "button[aria-label='Directions']"),
                (By.XPATH, "//button[@data-value='Directions']"),
            ],
            timeout=25,
        )
        directions_button.click()
        print("Clicked Directions.")

        start_input = wait_for_any_element(
            driver,
            [
                (By.CSS_SELECTOR, "input[aria-label^='Choose starting point']"),
                (By.CSS_SELECTOR, "input[aria-label^='Starting point']"),
                (By.CSS_SELECTOR, "input.tactile-searchbox-input[aria-label*='starting']"),
                (By.CSS_SELECTOR, "input.tactile-searchbox-input"),
            ],
            timeout=25,
        )
        start_input.clear()
        start_input.send_keys("Koramangala, Bangalore")
        start_input.send_keys(Keys.ENTER)
        print("Entered starting location.")

        destination_input = wait_for_any_element(
            driver,
            [
                (By.CSS_SELECTOR, "input[aria-label^='Choose destination']"),
                (By.CSS_SELECTOR, "input[aria-label^='Destination']"),
                (By.CSS_SELECTOR, "input.tactile-searchbox-input[aria-label*='destination']"),
                (By.CSS_SELECTOR, "input.tactile-searchbox-input"),
            ],
            timeout=25,
        )
        destination_input.clear()
        destination_input.send_keys("91 Springboard, Vikhroli")
        destination_input.send_keys(Keys.ENTER)
        print("Entered destination.")

        wait.until(EC.url_contains("/dir/"))

        route_cards = wait_for_any_elements(
            driver,
            [(By.CSS_SELECTOR, "div[data-trip-index]")],
            timeout=30,
        )
        if not route_cards:
            raise Exception("No routes were found for the provided locations.")

        route_cards[0].click()
        print("Selected first available route.")

        try_click_optional(
            driver,
            [
                (By.CSS_SELECTOR, "button[data-value='Steps']"),
                (By.CSS_SELECTOR, "button[aria-label='Steps']"),
                (By.CSS_SELECTOR, "button[data-value='Details']"),
                (By.CSS_SELECTOR, "button[aria-label='Details']"),
            ],
            timeout=5,
        )

        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[data-step-index]"))
            )
            step_elements = driver.find_elements(By.CSS_SELECTOR, "div[data-step-index]")
        except TimeoutException:
            print("No step elements found after waiting. Proceeding with empty list.")
            step_elements = []

        instructions = []
        for element in step_elements:
            text = element.text.strip()
            if not text:
                continue
            if text.lower() in {"directions", "steps"}:
                continue
            instructions.append(text)

        if not instructions:
            print("No step-by-step instructions were found.")
        else:
            print(f"Captured {len(instructions)} instructions.")

        base_dir = Path(__file__).resolve().parent
        excel_path = base_dir / "driving_instructions.xlsx"
        screenshot_path = base_dir / "route_screenshot.png"

        save_instructions_to_excel(instructions, str(excel_path))
        print("Saved instructions to driving_instructions.xlsx")

        save_full_page_screenshot(driver, str(screenshot_path))
        print("Saved screenshot to route_screenshot.png")

        print("Done.")
        return 0

    except TimeoutException:
        print("Timed out waiting for a page element. Please check selectors or network.")
        return 1
    except WebDriverException as exc:
        print(f"WebDriver error: {exc}")
        return 1
    except Exception as exc:
        print(f"Unexpected error: {exc}")
        return 1
    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    sys.exit(main())
