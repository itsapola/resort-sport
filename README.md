# Resort Sport Ltd.

*Official equipment for time off.*

The Collection 01 launch site for Resort Sport Ltd., a hypothetical, self-initiated brand case study. It puts into practice the three companion decks: Brand Strategy (Deck 01), Brand Identity & Guidelines (Deck 02) and Product & Go-to-Market (Deck 03).

| SKU   | Bag           | Type               | Proposed retail | Launch role (Deck 03) |
|-------|---------------|--------------------|-----------------|-----------------------|
| RS-01 | The Monaco    | Travel duffel      | $695–$895       | Arriving later        |
| RS-02 | The Rally     | Belt bag           | $275–$350       | Entry, available now  |
| RS-03 | The Riviera   | Oversized carryall | $595–$795       | Arriving later        |
| RS-04 | The Portofino | Handbag            | $495–$695       | Design hero, available now |
| RS-05 | The Regatta   | Backpack           | $525–$725       | In development        |

Every silhouette is designed in four permanent colorways: **Heritage** (natural canvas and cognac leather, the house signature), **Avorio** (ivory leather), **Cognac** and **Nero** (black leather). That makes 20 core variations. Every bag has a Sand lining, brushed silver-tone hardware and a tonal RS deboss; the Portofino adds a polished aluminum arch handle, under prototype testing.

## Launch assortment

Six SKUs can be added to the bag. Every other combination shows its expected price and a "Notify me" button that carries the request into the sign-up.

| Bag | Available now |
|---|---|
| The Portofino | Cognac, Avorio, Nero ($695) |
| The Rally | Heritage ($275), Cognac and Nero ($350) |

Pricing rule: Heritage sits at the low end of each range and the all-leather colorways at the high end. The launch Portofino is all leather, so it sells at $695, and the Deck 03 unit economics are modeled at $695 to match.

## What's on the page

1. Announcement bar and navigation, with a mobile menu
2. Hero: "Out of office. In our element." on Pool Aqua, with the Portofino in Cognac on a yacht deck as the cover photo
3. Collection 01: the Portofino and the Rally first and largest; the Monaco and Riviera arriving later; the Regatta in development
4. Campaign band: "Luxury doesn't have to be serious to be considered." (the creative provocation)
5. Manifesto: "The art of being off duty."
6. The Details: design worth recognizing
7. The four permanent colorways, shown on the Portofino
8. The Journal: three proposed stories
9. Limited Issues: Issue 001, the Rally in Curaçao, after launch
10. The Time Off Club: the correspondence sign-up
11. Footer

The product view follows Deck 03's Show, Explain, Reassure, Close: photos and colorways; the description; size, what fits, materials and build; care, delivery and returns; then price and the button. Portofino dimensions, capacity and weight are the Deck 03 design targets and are labeled as targets. Other bags state that their dimensions are published once prototypes are measured.

Microcopy carries the voice in small moments: "A fine choice" when something goes in the bag, "Nothing packed yet. Plans optional." when it's empty, and "Until next time." at checkout.

## Brand system in the code

- Line hierarchy: platform "The art of being off duty."; tagline "Official equipment for time off."; campaign "Out of office. In our element."; creative provocation "Luxury doesn't have to be serious to be considered." Other lines are section copy.
- Colorways: permanent colorways are named plainly; Limited Issues take cocktail names: Curaçao (Pool Aqua), Negroni (Bitter Orange), Limoncello (Butter Yellow) and Bijou (Racing Green).
- Type (Deck 02): Instrument Serif for display, in sentence case with selective italic; Space Grotesk Regular for body and Medium for labels.
- Color roles (Deck 02; see the `:root` tokens in `index.html`):
  - Cabana Red #811E23: brand authority, signature headlines, logo, primary actions
  - Deep Navy #11293A: text, navigation, editorial structure
  - Ivory #F1E9DD: primary canvas
  - Pool Aqua #98C6C5: resort accent (the cover) and the Curaçao Limited Issue
  - Sand #CFB9A2: the lining color; packaging neutral
  - Palm Green #123F34: supporting accent
  - Cognac, Avorio and Nero appear only as product colorway swatches
  - Panel tints are mixed from the palette, so there are no off-system colors

## Running it

The page is static HTML with no build step. Open `index.html` in a browser, or serve the folder:

```bash
python3 -m http.server
```

To host it on GitHub Pages, turn on Pages in the repo settings and choose the `main` branch, root folder. The link-preview image is addressed at the Pages URL.

## Notes

- Prices are proposed retail ranges, not final MSRP.
- Returns: 30 days from delivery, unused and in original condition with the dust bag. Shown in the product view and the bag. The delivery policy is not set yet.
- The Portofino stays at $695 by decision, about $155 above the Polène Numéro Neuf Mini; the pricing story needs to justify the gap.
- Photography: all five bags have studio cutouts in all four colorways. The hero (Portofino on a yacht deck) and Journal 002 (Portofino at a Paris café) use new editorial photos. Other lifestyle and detail photos still come from the early brand boards, and some of them show earlier designs that Deck 01 rejected (the barrel duffel, the conventional tote, the flap backpack).
- Colorways without a photo show a "Photo to come" tag.
- Checkout, journal stories, Limited Issues and Our World aren't connected; their buttons show a short note.
