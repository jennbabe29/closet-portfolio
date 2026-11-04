# Closet Portfolio

A closet manager built like a financial portfolio, for fashion lovers with a type-A personality.

Closet Portfolio treats your closet the way an investor treats a portfolio. Every piece you buy is an asset with a cost basis. Every day you own it adds to its holding period. When it sells, you see the realized gain or loss after platform fees. It's for fashion lovers who buy, wear and resell, and want to know financial metrics of their fashion investments.

It runs on Google Sheets with a Google Apps Script (JavaScript) add-on. Everything is free except the optional AI email reading, which uses your own Anthropic API key.

## How it works

Items move through three stages, plus one exit for pieces you don't sell. Each is its own tab:

| Stage | Tab | What it tracks |
| --- | --- | --- |
| Holdings | `Assets` | Everything you own: brand, asset class, segment, retail price, cost basis, discount, holding period |
| Listed | `Asset for Sale` | Pieces currently listed for resale |
| Sold | `Sale Log` | Sale price, broker (Poshmark, eBay, The Real Real...), payout and realized gain or loss |
| Written off | `Recycle/Donate/Gift` | Pieces you donated, gave away or threw out, and where they went |

## Features

- **Automatic calculations.** Type a new item's description and its formulas fill in: Discount, Hold (days), Hold Duration, Net Loss/Gain. Existing values are never overwritten.
- **Poshmark payouts.** Sold Profit fills in at 80% of the sale price, or price minus $2.95 under $15.
- **Add asset.** One click adds a dated row to the top of Holdings.
- **Easier data entry.** New rows get today's date (you can change it). Asset Class and Segment are dropdowns from a `Categories` tab you can edit, and each row only offers the Segments of its Asset Class. Size is a dropdown of the values you already use. Retail, Cost and Discount always show as dollars.
- **List for sale.** One click moves the selected asset from Holdings to Listed.
- **Mark as sold.** One click asks for the sale price and broker, then moves the item from Listed to the Sale Log.
- **Recycle, donate or gift.** One click asks where the piece went, then moves it from Holdings to `Recycle/Donate/Gift`.
- **Read order emails with AI (optional).** Label an order confirmation `Closet/To add` in Gmail and Claude reads it. Each piece becomes a new holding with its brand, description, category, color, size, retail price and pre-tax cost. It can check every hour on its own.

## Setup

1. Make a Google Sheet with tabs named `Assets`, `Asset for Sale`, `Sale Log` and `Recycle/Donate/Gift` (column layout below). The script leaves any other tab alone.
2. Open **Extensions → Apps Script**, paste in `closet_automation.gs` and click **Save**.
3. Reload the sheet. A **Closet** menu appears. Approve the permission prompt the first time.
4. Optional: add on-sheet buttons with **Insert → Drawing**, then **⋮ → Assign script** and enter `addAsset`, `listForSale` or `markAsSold`.
5. Optional, for email import: get an API key at [platform.claude.com](https://platform.claude.com) (billed separately, per use), and add it with **Closet → Set API key**. It's saved in the script's private settings, never in the code. In Gmail, create a filter that applies the label `Closet/To add` to order confirmations, then click **Closet → Check order emails every hour** once. Only emails with that label are sent to Anthropic, the company that makes Claude.

### Column layout

- `Assets` / `Asset for Sale` / `Recycle/Donate/Gift`: A Date Acquired · B Brand · C Asset Class · D Segment · E Description · F Color · G Size · H Seller · I Retail · J Cost · K Discount · L Hold · M Hold Duration · N Notes
- `Sale Log`: A Date Sold · B Date Acquired · C Brand · D Asset Class · E Segment · F Description · G Color · H Size · I Seller · J Retail · K Cost · L Discount · M Sold Price · N Shipping · O Notes · P Sold Profit · Q Net Loss/Gain · R Broker · S Held · T Hold

## Note

The buttons and menu work in Google Sheets on a computer. The Sheets mobile app doesn't run Apps Script menus or buttons.

## License

MIT
