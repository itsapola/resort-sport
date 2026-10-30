# Resort Sport Ltd.

*Official equipment for time off.*

A one-page storefront prototype for Resort Sport Ltd. that shows the five-piece collection:

| Code  | Piece         | Use                    |
|-------|---------------|------------------------|
| RS-01 | The Duffel    | For the weekend        |
| RS-02 | The Racer     | For every day          |
| RS-03 | The Riviera   | For the beach          |
| RS-04 | The Portofino | Handbag                |
| RS-05 | The Regatta   | Backpack               |

## What's on the page

- Product cards with Ivory, Cognac and Black colorways, plus the limited-edition Oxblood Portofino
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

- Prices, specs and product copy are placeholders.
- Cormorant Garamond stands in for Canela, which isn't available as a free web font.
- The Duffel and Portofino use studio photography. The Racer, Riviera and Regatta still use crops from the brand boards.
- Checkout and email sign-up don't send anything anywhere yet.
