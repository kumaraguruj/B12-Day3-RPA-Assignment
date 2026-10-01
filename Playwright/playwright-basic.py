from playwright.sync_api import sync_playwright
from pathlib import Path

DEFAULT_URL = (
    "https://www.cricbuzz.com/live-cricket-scorecard/151543/"
    "ind-vs-wi-2nd-odi-west-indies-tour-of-india-2026"
)

# Start the Playwright automation session.
print("Starting Playwright...")
with sync_playwright() as p:
    # Launch Chromium and open a new browser page.
    print("Launching Chromium...")
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    print("Browser page created.")

    # Open the requested India versus West Indies scorecard.
    print("Opening the Cricbuzz scorecard...")
    page.goto(DEFAULT_URL, wait_until="domcontentloaded", timeout=60_000)
    # Wait for both innings scorecards, then extract their text.
    print("Waiting for scorecard innings...")
    innings_sections = page.locator("div[id^='scard-team-'][id*='-innings-']")
    innings_sections.first.wait_for(state="visible", timeout=20_000)
    innings_data = innings_sections.all_inner_texts()
    sections = []
    for index in range(innings_sections.count()):
        section = innings_sections.nth(index)
        text = section.inner_text().replace("\u00a0", " ")
        text = "\n".join(line.strip() for line in text.splitlines() if line.strip())
        if "Batter" in text and "Bowler" in text:
            sections.append(text)

    # Add the match title and innings data to the text output.
    title = page.locator("h1").first
    header = title.inner_text().strip() if title.count() else "India vs West Indies scorecard"
    match_data = f"{header}\n\n" + "\n\n".join(sections)

    # Save the scorecard text next to this script.
    output_path = Path(__file__).resolve().parent / "match_data.txt"
    output_path.write_text(match_data + "\n", encoding="utf-8")
    print(f"Saved scorecard text to {output_path}")

    # Close the browser when scraping is complete.
    print("Closing the browser...")
    browser.close()
    print("Browser closed.")