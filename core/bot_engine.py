"""
RPA Bot Engine - Core automation module
FIXES APPLIED:
  1. click_submit()        — no longer raises after successful navigation;
                             detects success via URL change, not click confirmation
  2. fill_form()           — verifies final page URL to determine SUCCESS/FAILED
  3. _submit_succeeded()   — 3-tier check: URL change, config fragment, DOM selector
  4. _normalize_time()     — converts '08:00 PM' → '20:00' for <input type='time'>
  5. fill_time()           — auto-normalises value before sending
  6. select_radio()        — exact match only, no greedy partial match
  7. screenshot()          — reports/ dir created eagerly in __init__
  8. fill_form()           — required empty fields now RAISE (not just warn), blocking submit
  9. select_checkboxes()   — bidirectional fuzzy match so 'Extra Cheese' matches value='cheese'
                             and value='extracheese' matches label='Extra Cheese'
"""

import time
import logging
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementNotInteractableException, ElementClickInterceptedException,
)

logger = logging.getLogger(__name__)

BY_MAP = {
    "id":    By.ID,
    "css":   By.CSS_SELECTOR,
    "xpath": By.XPATH,
    "name":  By.NAME,
    "tag":   By.TAG_NAME,
    "class": By.CLASS_NAME,
}

def _resolve_by(sel):
    return BY_MAP[sel.get("by", "css").lower()], sel["value"]

def _token(s: str) -> str:
    """Normalise a string for fuzzy comparison: lowercase, strip spaces/punctuation."""
    import re
    return re.sub(r"[^a-z0-9]", "", s.strip().lower())


class RPABotEngine:

    def __init__(self, config, headless=False):
        self.config   = config
        self.headless = headless
        self.driver   = None
        self.wait     = None
        self.results  = []
        Path("reports").mkdir(exist_ok=True)

    # ── Driver ────────────────────────────────────────────────────────────────

    def start_driver(self):
        opts = Options()
        if self.headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--log-level=3")
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.driver = webdriver.Chrome(options=opts)
        self.wait   = WebDriverWait(self.driver, self.config.get("timeout", 15))
        logger.info("✅ Chrome WebDriver started")

    def stop_driver(self):
        if self.driver:
            self.driver.quit()
            logger.info("🛑 Browser closed")

    # ── Navigation ────────────────────────────────────────────────────────────

    def navigate(self, url):
        self.driver.get(url)
        WebDriverWait(self.driver, 15).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(self.config.get("page_load_delay", 1.0))
        logger.info(f"🌐 Navigated to: {url}")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _scroll(self, el):
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center',inline:'nearest'});", el
        )
        time.sleep(0.15)

    def _js_click(self, el):
        self.driver.execute_script("arguments[0].click();", el)

    def _js_set(self, el, value):
        self.driver.execute_script(
            """
            arguments[0].focus();
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input',  {bubbles:true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
            """, el, value
        )

    def find_el(self, selector):
        return self.wait.until(EC.visibility_of_element_located(_resolve_by(selector)))

    def _label(self, el):
        eid = el.get_attribute("id")
        if eid:
            try:
                return self.driver.find_element(
                    By.CSS_SELECTOR, f"label[for='{eid}']"
                ).text.strip()
            except NoSuchElementException:
                pass
        try:
            return el.find_element(By.XPATH, "./ancestor::label[1]").text.strip()
        except NoSuchElementException:
            pass
        return el.get_attribute("aria-label") or el.get_attribute("title") or ""

    @staticmethod
    def _normalize_time(value: str) -> str:
        """Convert any reasonable time string to HH:MM (24-hour)."""
        v = value.strip()
        for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M", "%H:%M:%S"):
            try:
                return datetime.strptime(v, fmt).strftime("%H:%M")
            except ValueError:
                continue
        logger.warning(f"⚠️  Could not normalise time value '{v}' — sending as-is")
        return v

    # ── Field handlers ────────────────────────────────────────────────────────

    def fill_text(self, selector, value):
        el = self.find_el(selector)
        self._scroll(el)
        self._js_set(el, str(value) if value is not None else "")
        logger.debug(f"✏️  text  '{selector['value']}' = '{value}'")

    def fill_time(self, selector, value):
        raw = str(value).strip()
        if not raw:
            return
        normalised = self._normalize_time(raw)
        el = self.find_el(selector)
        self._scroll(el)
        self._js_set(el, normalised)
        logger.debug(f"⏰ time  '{selector['value']}' = '{normalised}' (raw: '{raw}')")

    def select_dropdown(self, selector, value):
        el  = self.find_el(selector)
        self._scroll(el)
        sel = Select(el)
        v   = str(value).strip()
        for fn in [
            lambda: sel.select_by_visible_text(v),
            lambda: sel.select_by_value(v),
            lambda: sel.select_by_value(v.lower()),
            lambda: [o.click() for o in sel.options if v.lower() in o.text.lower()][0],
        ]:
            try:
                fn(); return
            except Exception:
                continue
        raise RuntimeError(f"Dropdown: cannot select '{v}' in {selector}")

    def select_radio(self, name, value):
        target = value.strip().lower()
        radios = self.driver.find_elements(
            By.CSS_SELECTOR, f"input[type='radio'][name='{name}']"
        )
        if not radios:
            raise NoSuchElementException(f"No radios with name='{name}'")
        for r in radios:
            rv = (r.get_attribute("value") or "").strip().lower()
            lv = self._label(r).strip().lower()
            if target == rv or target == lv:
                self._scroll(r)
                self._js_click(r)
                logger.debug(f"🔘 radio name='{name}' → '{rv}' (matched '{value}')")
                return
        avail = [
            f"value='{r.get_attribute('value')}' label='{self._label(r)}'"
            for r in radios
        ]
        raise ValueError(f"Radio '{name}': no match for '{value}'. Available: {avail}")

    def select_checkboxes(self, name, values):
        """
        FIX 9: Bidirectional fuzzy match using tokenised strings.

        Problem: CSV has 'Extra Cheese', form has value='cheese' and label='Extra Cheese'.
        Old code: `cv in wanted` checks if 'cheese' is in {'extra cheese'} → False ❌
        Fix: also check if _token(cv) matches _token(wanted_item) and vice versa,
             so 'cheese' matches 'Extra Cheese' and 'extracheese' matches 'extra cheese'.
        """
        if isinstance(values, list):
            raw_wanted = [v.strip() for v in values if str(v).strip()]
        else:
            raw_wanted = [v.strip() for v in str(values).split(",") if v.strip()]

        # Build two lookup sets: exact lowercase and tokenised (no spaces/punctuation)
        wanted_lower  = {v.lower() for v in raw_wanted}
        wanted_tokens = {_token(v) for v in raw_wanted}

        cbs = self.driver.find_elements(
            By.CSS_SELECTOR, f"input[type='checkbox'][name='{name}']"
        )
        if not cbs:
            raise NoSuchElementException(f"No checkboxes with name='{name}'")

        matched = set()
        for cb in cbs:
            cv  = (cb.get_attribute("value") or "").strip()
            lv  = self._label(cb).strip()
            cv_l, cv_t = cv.lower(), _token(cv)
            lv_l, lv_t = lv.lower(), _token(lv)

            # FIX: check all four combinations so 'Extra Cheese' ↔ 'cheese' both work
            hit = (
                cv_l in wanted_lower or lv_l in wanted_lower
                or cv_t in wanted_tokens or lv_t in wanted_tokens
            )

            if hit and not cb.is_selected():
                self._scroll(cb)
                self._js_click(cb)
                matched.add(cv_l)
                logger.debug(f"☑️  checkbox '{name}' checked '{cv}' (label='{lv}')")
            elif not hit and cb.is_selected():
                self._scroll(cb)
                self._js_click(cb)
                logger.debug(f"⬜  checkbox '{name}' unchecked '{cv}'")

        unmatched = wanted_lower - matched
        if unmatched:
            avail = [
                f"value='{c.get_attribute('value')}' label='{self._label(c)}'"
                for c in cbs
            ]
            logger.warning(
                f"⚠️  Checkbox '{name}': values {unmatched} not matched. Available: {avail}"
            )

    def click_element(self, selector):
        el = self.find_el(selector)
        self._scroll(el)
        try:
            el.click()
        except ElementClickInterceptedException:
            self._js_click(el)
        logger.debug(f"🖱️  clicked '{selector['value']}'")

    # ── Submit ────────────────────────────────────────────────────────────────

    def click_submit(self, selector):
        """
        Clicks submit and waits for page to change.
        Does NOT raise on navigation — _submit_succeeded() makes the final call.
        """
        submit_wait = self.config.get("submit_wait", 8)
        url_before  = self.driver.current_url

        elements = self.driver.find_elements(*_resolve_by(selector))
        if not elements:
            raise NoSuchElementException(
                f"No submit button found for selector: {selector}"
            )

        found_texts = [
            e.text.strip() or e.get_attribute("value") or "(no text)"
            for e in elements
        ]
        logger.debug(f"🔍 Submit candidates: {found_texts}")

        for el in elements:
            try:
                self._scroll(el)
                if not (el.is_displayed() and el.is_enabled()):
                    continue
                try:
                    el.click()
                except (ElementClickInterceptedException, ElementNotInteractableException):
                    self._js_click(el)
                logger.info(
                    f"🚀 Submit clicked: '{el.text.strip() or el.get_attribute('value')}'"
                )
                break
            except Exception as e:
                # StaleElementReferenceException = page already navigated = success
                logger.debug(f"⏭️  Submit candidate skipped ({type(e).__name__}): {e}")
                break

        try:
            WebDriverWait(self.driver, submit_wait).until(
                lambda d: d.current_url != url_before
                or d.execute_script("return document.readyState") == "complete"
            )
        except TimeoutException:
            logger.warning("⚠️  Page did not change within submit_wait timeout")

        time.sleep(self.config.get("submit_delay", 1.5))

        new_url = self.driver.current_url
        if new_url != url_before:
            logger.info(f"✅ URL changed after submit → {new_url}")
        else:
            logger.warning(f"⚠️  URL unchanged after submit ({new_url})")

    def _submit_succeeded(self, url_before: str) -> bool:
        """
        3-tier success check:
          Tier 1: any URL change
          Tier 2: config 'success_url_contains' fragment in new URL
          Tier 3: config 'success_selector' DOM element present
        """
        current = self.driver.current_url

        if current != url_before:
            logger.debug(f" Success: URL changed → {current}")
            return True

        fragment = self.config.get("success_url_contains", "")
        if fragment and fragment in current:
            logger.debug(f"Success: URL contains '{fragment}'")
            return True

        sel = self.config.get("success_selector")
        if sel:
            try:
                self.driver.find_element(*_resolve_by(sel))
                logger.debug(" Success: success_selector found in DOM")
                return True
            except NoSuchElementException:
                pass

        logger.warning(f"⚠️  Still on {current} after submit — marking FAILED")
        return False

    # ── Form fill ─────────────────────────────────────────────────────────────

    def fill_form(self, row, field_map):
        url_before = self.driver.current_url
        result = {
            "row_data":  row,
            "status":    "SUCCESS",
            "error":     None,
            "timestamp": datetime.now().isoformat(),
        }
        try:
            for field in field_map:
                col   = field["column"]
                ftype = field.get("type", "text").lower()
                value = row.get(col, "")

                if str(value).strip() == "":
                    if field.get("required", False):
                        # FIX 8: raise instead of just warn — empty required field
                        # must block the submit, not silently produce a broken form
                        raise ValueError(
                            f"Required field '{col}' is empty — cannot submit form"
                        )
                    else:
                        logger.debug(f"⏭️  skip empty '{col}'")
                    if not field.get("fill_empty", False):
                        continue

                if   ftype == "text":     self.fill_text(field["selector"], value)
                elif ftype == "time":     self.fill_time(field["selector"], value)
                elif ftype == "dropdown": self.select_dropdown(field["selector"], value)
                elif ftype == "radio":
                    rn = field.get("radio_name")
                    if not rn:
                        raise ValueError(f"Field '{col}' has type 'radio' but no 'radio_name'")
                    self.select_radio(rn, value)
                elif ftype == "checkbox":
                    cn = field.get("checkbox_name")
                    if not cn:
                        raise ValueError(f"Field '{col}' has type 'checkbox' but no 'checkbox_name'")
                    self.select_checkboxes(cn, value)
                elif ftype == "click":    self.click_element(field["selector"])
                else:
                    logger.warning(f"Unknown field type '{ftype}' for column '{col}' — skipped")

            if "submit_selector" in self.config:
                self.click_submit(self.config["submit_selector"])

                if not self._submit_succeeded(url_before):
                    result["status"] = "FAILED"
                    result["error"]  = (
                        "Submit clicked but page did not change — "
                        "possible form validation error or wrong submit selector"
                    )
                    try:
                        self.screenshot(f"failed_{datetime.now().strftime('%H%M%S%f')}.png")
                    except Exception:
                        pass

        except Exception as exc:
            result["status"] = "FAILED"
            result["error"]  = str(exc)
            logger.error(f"❌ {exc}")
            try:
                self.screenshot(f"error_{datetime.now().strftime('%H%M%S%f')}.png")
            except Exception:
                pass

        return result

    # ── Batch ─────────────────────────────────────────────────────────────────

    def process_batch(self, records, field_map):
        url = self.config["form_url"]
        self.results = []
        for idx, row in enumerate(records, 1):
            logger.info(f"📋 Processing record {idx}/{len(records)}")
            try:
                self.navigate(url)
                result = self.fill_form(row, field_map)
            except Exception as exc:
                result = {
                    "row_data":  row,
                    "status":    "ERROR",
                    "error":     str(exc),
                    "timestamp": datetime.now().isoformat(),
                }
                logger.exception(f"💥 Record {idx} unexpected error")
            self.results.append(result)
            icon = "✅" if result["status"] == "SUCCESS" else "❌"
            logger.info(
                f"{icon} Record {idx} → {result['status']}"
                + (f" | {str(result['error'])[:120]}" if result["error"] else "")
            )
            time.sleep(self.config.get("delay_between_records", 1))

        ok = sum(1 for r in self.results if r["status"] == "SUCCESS")
        logger.info(f"🏁 Batch complete: {ok}/{len(self.results)} succeeded")
        return self.results

    # ── Screenshot ────────────────────────────────────────────────────────────

    def screenshot(self, filename=None):
        if not filename:
            filename = f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        path = Path("reports") / filename
        path.parent.mkdir(exist_ok=True)
        self.driver.save_screenshot(str(path))
        logger.info(f"📸 {path}")
        return str(path)