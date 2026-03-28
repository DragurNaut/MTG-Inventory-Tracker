# Requirements
Requires JSONConverter.bas
This should be preloaded, if not, get it here and import it into the code editor as a module 
https://github.com/VBA-tools/VBA-JSON

# MTG Set Tracker (Excel + Scryfall API)

An Excel-based Magic: The Gathering set tracker powered by the [Scryfall API](https://scryfall.com/docs/api).

This workbook is designed to help you quickly load a full MTG set into Excel, track card information, update market prices, and visually organize your collection with color-coded formatting.

---

## Features

* Load full MTG set data using a **set code**
* Pull card details directly from **Scryfall**
* Track:

  * Set Code
  * Collector Number
  * Card Name
  * Mana Cost
  * Color Identity
  * Type Line
  * Rarity
  * USD Price
  * USD Foil Price
* Special handling for:

  * Double-faced / multi-faced cards
  * Color identity styling
  * Rarity styling
  * Price highlighting
* Reusable workbook template for loading new sets

---

## Data Source

This project uses the public [Scryfall API](https://scryfall.com/docs/api) for card and pricing data.

Scryfall is an excellent MTG data source and deserves the credit here.

---

## Requirements

* Microsoft Excel (desktop version)
* Macro support enabled
* Internet connection (required for API calls)

> **Note:** Excel Online / browser versions generally do **not** support VBA macros properly.

---

## Getting Started

### 1) Download the Workbook

Download the `.xlsm` workbook from this repository.

### 2) Unblock the File (Important)

Because this workbook is downloaded from the internet, Windows may block macros automatically.

Before opening the workbook:

1. Right-click the `.xlsm` file
2. Click **Properties**
3. At the bottom, check **Unblock**
4. Click **Apply**
5. Click **OK**

### 3) Open in Excel

Open the workbook in **Microsoft Excel (desktop)**.

If prompted:

* Click **Enable Editing**
* Click **Enable Content** / **Enable Macros**

---

## Macro / Security Notes

This workbook contains VBA macros used to:

* Load card data from Scryfall
* Update pricing data
* Apply formatting and styling

It does **not**:

* install software
* change Windows settings
* access unrelated files
* run shell / command prompt scripts

If you want to inspect the VBA code before enabling macros:

* Open the workbook in Excel
* Press `Alt + F11`

---

## Recommended Excel Setup

Instead of lowering macro security globally, it is better to store this workbook in a trusted folder.

### Add a Trusted Location

In Excel:

1. Go to **File → Options**
2. Open **Trust Center**
3. Click **Trust Center Settings**
4. Open **Trusted Locations**
5. Add the folder where you keep this workbook

This allows Excel to trust files in that folder without reducing overall security.

---

## How to Use

## Load a Set

1. Open the workbook
2. Enter the **set code** into the input cell on the template sheet

   * Example:

     * `blb`
     * `mh3`
     * `spm`
3. Click the **Load Set** button

This will populate the sheet with card data from that set.

---

## Update Prices

Price updating is kept separate from the main set load.

1. Load or prepare your sheet
2. Click the **Update Prices** button

This will refresh:

* **USD Price**
* **USD Foil Price**

based on the current data from Scryfall.

---

## Included Formatting

The workbook includes visual formatting to make scanning easier:

### Color Identity Styling

* Single-color cards are color-coded by identity
* Multi-color cards are styled with separated color symbols

### Rarity Styling

* Common
* Uncommon
* Rare
* Mythic

### Price Highlighting

Price columns are color-coded by value range.

Example grouping:

* **$2–$5**
* **$6–$10**
* **$11–$15**
* **$16–$20**
* **$20+**

Foil pricing is highlighted separately to stand out more clearly.

---

## Known Limitations

* Some cards may not have:

  * normal USD price
  * foil USD price
  * special finish pricing (etched, surge foil, etc.)
* Certain promo / alternate printings may behave differently depending on Scryfall data
* Not every set includes every finish type
* Excel macro behavior can vary slightly depending on version / security settings

---

## Project Structure

This workbook is organized around a few main VBA components:

* **Set Loader**

  * Pulls full set data from Scryfall
* **Price Updater**

  * Refreshes market prices separately
* **Styling Helpers**

  * Handles card color identity, rarity, and price formatting

---

## Why This Exists

I wanted a reusable Excel-based MTG set tracker that was:

* lightweight
* customizable
* visually useful
* not locked behind a paid collection tool

This project is built for MTG players, collectors, spreadsheet nerds, and anyone who likes having control over their own data.

---

## Roadmap / Future Ideas

Possible future improvements:

* TCGplayer market price support
* Card image / Scryfall link support
* Pull by collector number or card name
* Better promo / finish handling
* Collection ownership / quantity tracking
* Sorting / filtering helpers
* Automatic conditional formatting setup
* Better template sheet UX

---

## Contributing

If you want to improve the workbook, feel free to:

* fork the repo
* suggest fixes
* improve VBA structure
* add quality-of-life features

If you find an issue or edge case, open an issue or submit a pull request.

---

## Credits

* [Scryfall](https://scryfall.com/) for the API and card data
* Microsoft Excel for being weirdly perfect for this kind of project

---

## License

You can add your preferred license here.

Examples:

* MIT
* GPL
* Personal use only

If you are not sure, MIT is a solid simple choice for a project like this.

---

## Screenshots

Add screenshots here if you want to show:

* Template sheet
* Loaded set example
* Styled rarity / color identity
* Price highlighting

Example:

```md
![Template Screenshot](images/template.png)
```

---

## Final Note

If the workbook appears to “do nothing,” macros are almost always the reason.

Make sure the file is:

* **unblocked**
* opened in **desktop Excel**
* macros are **enabled**

Then it should work as intended.
