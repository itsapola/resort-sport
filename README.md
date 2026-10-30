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

- Product cards with the four permanent colorways, plus the limited-edition Oxblood Portofino
- Studio cutout photos (`img/*-studio-*.webp`) for the Duffel and Portofino
- A quick view for each product, with photos, colorway picker and specs
- A bag that keeps items in the browser's local storage
- A sign-up list for limited-edition colors

## Running it

The page is static HTML with no build step. Open `index.html` in a browser, or serve the folder:

```bash
python3 -m http.server
```

To host it on GitHub Pages, turn on Pages in the repo settings and choose the `main` branch, root folder.

## Notes

- Prices are proposed retail ranges, not final MSRP. Until per-variant pricing is set, the page charges the low end for Heritage and the high end for the all-leather colorways (and for Oxblood).
- Dimensions and product copy are placeholders.
- Colorways without a photo yet show the Heritage photo with a "Photo to come" tag.
- Cormorant Garamond stands in for Canela, which isn't available as a free web font.
- The Duffel and Portofino use studio photography. The Racer, Riviera and Regatta still use crops from the brand boards.
- Checkout and email sign-up don't send anything anywhere yet.
