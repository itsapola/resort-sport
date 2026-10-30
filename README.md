# Resort Sport Ltd.

*Official equipment for time off.*

A one-page storefront prototype for Resort Sport Ltd. that shows the five-piece collection:

| SKU   | Bag           | Type               | Proposed retail |
|-------|---------------|--------------------|-----------------|
| RS-01 | The Monaco    | Travel duffel      | $695–$895       |
| RS-02 | The Rally     | Belt bag           | $275–$350       |
| RS-03 | The Riviera   | Oversized carryall | $595–$795       |
| RS-04 | The Portofino | Handbag            | $495–$695       |
| RS-05 | The Regatta   | Backpack           | $525–$725       |

Every bag comes in four permanent colorways: **Heritage** (ivory canvas and cognac leather), **Cognac**, **Latte** (cream leather) and **Noir** (black leather). That makes 20 core variations. All bags have a sand-colored lining, silver-tone hardware, a tonal RS monogram and protected leather or canvas finishes.

## What's on the page

The homepage follows the V1 editorial copy, in this order:

1. Announcement bar and navigation
2. Hero: "Out of office. In our element." on a sun-yellow cover
3. The Permanent Collection: an editorial grid with the Monaco and Portofino in the largest slots and a full-width Regatta card
4. "The art of being off duty" manifesto
5. The Details: leather, hardware and close-ups
6. The four colorways, shown on the Portofino
7. The Journal: three proposed stories
8. Limited Issues: the Rally in Pool Aqua, shown as a concept
9. Correspondence: the email sign-up
10. Footer

Each product opens a quick view with colorways, price and specs. The bag keeps items in the browser's local storage and uses the brand's microcopy ("A fine choice", "Until next time").

## Running it

The page is static HTML with no build step. Open `index.html` in a browser, or serve the folder:

```bash
python3 -m http.server
```

To host it on GitHub Pages, turn on Pages in the repo settings and choose the `main` branch, root folder.

## Notes

- Prices are proposed retail ranges, not final MSRP. Until per-variant pricing is set, the page charges the low end for Heritage and the high end for the all-leather colorways.
- Dimensions and product copy are placeholders.
- Colorways without a photo yet show the Heritage photo with a "Photo to come" tag.
- Cormorant Garamond stands in for Canela, which isn't available as a free web font.
- The Duffel and Portofino use studio photography. The Racer, Riviera and Regatta still use crops from the brand boards.
- Checkout, journal stories, Limited Issues and Our World aren't connected yet; their buttons show a short note.
- The hero is a placeholder layout built from the existing photography. It should be replaced by the campaign shoot (speedboat, terrace or convertible) when that exists.
