---
name: ExcelKey
colors:
  background: '#ffffff'
  on-background: '#000000'
  surface: '#ffffff'
  surface-container: '#f4f4f4'
  surface-container-low: '#f2f0ed'
  on-surface: '#0d0d0d'
  on-surface-variant: '#374151'
  outline: '#c8c8ce'
  outline-variant: '#e5e5e5'
  primary: '#1377f9'
  on-primary: '#ffffff'
  primary-hover: '#0e63d6'
  primary-container: '#eaf2fe'
  on-primary-container: '#000000'
  link-hover: '#4d9bff'
  secondary: '#374151'
  on-secondary: '#ffffff'
  error: '#f91313'
  warning: '#e99a32'
  dark-background: '#0c0c0e'
  dark-surface: '#161618'
  dark-surface-container: '#1f1f23'
  dark-surface-hover: '#26262a'
  dark-on-surface: '#f2f2f5'
  dark-on-surface-variant: '#c9ced6'
  dark-outline: '#6a6a70'
  dark-primary-container: '#16263f'
  dark-link: '#4d9bff'
  dark-link-hover: '#7db4ff'
  dark-error: '#ff5a52'
  dark-warning: '#f0b86a'
typography:
  wordmark:
    fontFamily: Inter
    fontSize: 26px
    fontWeight: '400'
    lineHeight: normal
    letterSpacing: '0'
  heading:
    fontFamily: Inter
    fontSize: 17px
    fontWeight: '600'
    lineHeight: 1.45
    letterSpacing: '0'
  body:
    fontFamily: Inter
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 1.45
    letterSpacing: '0'
  button:
    fontFamily: Inter
    fontSize: 15px
    fontWeight: '600'
    lineHeight: '1'
    letterSpacing: '0'
  auth-button:
    fontFamily: Inter
    fontSize: 14px
    fontWeight: '600'
    lineHeight: '1'
    letterSpacing: '0'
rounded:
  sm: 8px
  DEFAULT: 10px
  lg: 14px
  full: 999px
spacing:
  small: 8px
  control-gap: 10px
  card-padding: 16px
  message-margin: 22px
  action-gap: 24px
  avatar-indent: 50px
  mobile-gutter: 5vw
---

# Design System: ExcelKey

Source: [filekey.app](https://filekey.app/), inspected October 3, 2026; the live menu reports v1.12.1. This ExcelKey design system adapts the visual system extracted from the original page's embedded stylesheet, CSS custom properties, rendered DOM, and desktop/mobile visual inspection. The local project's modified `web/index.html` is not the authority for the source extraction.

Scope: the core file-encryption application. Authenticated components are documented from shipped CSS, rather than an authenticated session. Blog and legal-page variants are outside this extraction. YAML uses light-theme defaults; prefixed dark tokens describe the alternate theme. `outline-variant` is an opaque approximation of the actual translucent black border on white, not a native CSS variable.

## 1. Visual Theme & Atmosphere

ExcelKey feels like a quiet, helpful utility. A clean white canvas, near-black prose, and a precise electric-blue accent give it a direct, dependable character. The interface begins with conversational instructions and compact actions; generous empty space allows the task to remain the focal point. Rounded controls and small lock emblems soften the technical subject without decorative illustration or a large display headline.

Dark mode preserves that same visual rhythm on a near-black canvas, with charcoal surfaces and brighter blue links. Elevation is restrained: most content sits flat on the page, while small floating menus receive a fine border and soft shadow. The hierarchy comes from weight, alignment, spacing, and role-specific color. The design reads as a conversation about files, with system messages on the left and uploaded items aligned toward the right.

## 2. Color Palette & Roles

### Primary Foundation

| Descriptive name | Light | Dark counterpart | Role |
| --- | --- | --- | --- |
| Paper White / Midnight Ink | `#ffffff` | `#0c0c0e` | Page, fixed header band, bottom dock wrapper |
| Clean White / Soft Charcoal | `#ffffff` | `#161618` | Menus, inputs, download surfaces; dark `--fk-surface` |
| Mist Gray / Graphite | `#f4f4f4` | `#1f1f23` | Uploaded cards, contact chips, filled technical blocks; `--fk-fill` |
| Warm Pebble | `#f2f0ed` | `#26262a` | Secondary icon-tile background; merges into raised charcoal in dark mode |
| Pale Sky / Deep Blue Slate | `#eaf2fe` | `#16263f` | Bottom file-action dock |
| Raised Charcoal | — | `#26262a` | Dark icon tiles and ghost-button hover |

### Accent & Interactive

| Descriptive name | Hex | Role |
| --- | --- | --- |
| ExcelKey Electric Blue | `#1377f9` | Primary button, lock/file icons, light links, active appearance option |
| Pressed Cobalt | `#0e63d6` | Primary-button hover in both themes |
| Bright Link Blue | `#4d9bff` | Light-link hover and default dark links |
| Icy Highlight Blue | `#7db4ff` | Dark-link hover |

Primary buttons retain electric blue and white labels in dark mode. Blue indicates an available action. Inline links have no underline; external links append a small arrow icon whose stroke follows the text color.

### Typography & Text Hierarchy

| Descriptive name | Light | Dark | Role |
| --- | --- | --- | --- |
| True Ink / Frost White | `#000000` | `#f2f2f5` | Conversation text and wordmark |
| Soft Ink / Frost White | `#0d0d0d` | `#f2f2f5` | Form text, names, ghost-button text; `--fk-ink-soft` |
| Slate Text / Silver Slate | `#374151` | `#c9ced6` | Menu entries and inactive appearance choices |
| Quiet Silver / Ash Gray | `#9ca3af` | `#6a6a70` | Version label; dark import icon also uses ash gray |
| Pale Silver | `#c8c8ce` | `#6a6a70` | Import icon stroke |

Muted text is primarily opacity-based: `#00000080` for light status text (about 50% black), `#00000099` for the stronger muted token (60%), and `#0000003a` for faint details (23%). Dark status text uses `#ffffff66` (40%); the stronger muted token is `#ffffff80`, and the faint token is `#ffffff3a`. Preserve alpha rather than treating all gray roles as one opaque color.

### Functional States

- **Alert Red `#f91313` / Coral Alert `#ff5a52`:** error lock icon and destructive hover; `--fk-error` switches with theme.
- **Amber Warning `#e99a32` / Warm Gold `#f0b86a`:** warning lock icon.
- **Electric Blue `#1377f9`:** normal informational lock icon. The core stylesheet defines no dedicated green success palette.
- Status-avatar borders use the matching color at roughly 13% opacity in light mode (`21` hex alpha) and 33% in dark mode (`55`).
- General borders use `#0000001a` in light mode and `#ffffff1f` in dark mode. Ghost-button borders use `#0000001f` / `#ffffff2b`; identity dividers use `#0000000d` / `#ffffff14`.
- Drag targeting uses a translucent aqua overlay, `#c0f0ff33`. Contact hover uses translucent electric blue: `#1377f91a` in light mode and `#1377f933` in dark mode.

## 3. Typography Rules

Use self-hosted **Inter**, a variable sans-serif supplied by `/fonts/inter.woff2`, with `system-ui` fallback. Its clear, neutral shapes support both instructions and technical material. The font file supports weights 100–900; the interface uses a small practical subset. Menu section captions explicitly use `system-ui`. Technical fingerprints and keys use `ui-monospace`, `SFMono-Regular`, `Menlo`, then `monospace`.

### Hierarchy & Weights

| Role | Size | Weight | Line height / tracking |
| --- | --- | --- | --- |
| Wordmark | 26px | 400 inherited | Normal; default tracking |
| Intro emphasis | 16px | 700 (`strong`) | 1.45 / 23.2px |
| Conversation body | 16px | 400 | 1.45 / 23.2px |
| Message section heading | 17px | 600 | Inherits message rhythm |
| Numbered message heading | 16px | 600 | Inherits message rhythm |
| File title | 16px inherited | 600 | 1.0 |
| Dock button | 15px | 600 primary, 500 ghost | 1.0 |
| First-run button | 14px | 600 primary, 500 ghost | 1.0 |
| Inline action | 16px inherited | 450 links; 500 icon actions | Prose rhythm or 1.0 for action rows |
| Menu item | 15px | 420 | 1.25 |
| Menu section caption | 11px, system-ui | 600 | Normal; visually uppercase |
| Account identity name | 14px | 600 | 1.35 |
| Account fingerprint | 12px, monospace | 400 | Normal; 0.03em tracking |
| Contact name / chip | 15px / 14px | 600 / 400 | 1.2 / 1.0 |
| Contact key | 11px, monospace | 400 | 1.4 |
| Picker helper / empty state | 13px / 14px | 400 | 1.4 / 1.5 |
| Version label | 11px | 250 | Normal |
| Notification banner | 20px; 16px on small portrait screens | 500 | 20px |

There is no conventional oversized H1–H6 ladder in the core app. Its opening headline is bold body text embedded in the first message. Do not substitute a marketing-style hero for this compact hierarchy.

### Spacing Principles

Body copy is comfortably readable at 145% line height, while titles and action labels are compact. Tracking is generally the browser default; extra tracking is reserved for fingerprints. Status values use tabular numerals. Message wrappers carry 22px top and bottom margins; forms and action rows use smaller 8–15px offsets. Weight supplies emphasis without enlarging every heading.

## 4. Component Stylings

### Buttons

Primary buttons have electric-blue fill, white text, 10px corners, a transparent 1px border, and 11px vertical by 18px horizontal padding. Ghost buttons have white fill, soft-ink labels, a faint border, and medium weight; their hover fill is mist gray. Dark ghosts use graphite fill and raised-charcoal hover.

First-run Unlock/Create buttons use 14px labels, 8px by 14px padding, 15px icons, and a 7px icon gap. Their row wraps with an 8px gap and 12px top margin. On hover, these buttons lift 1px and gain a small shadow; the fingerprint scales to 1.12 and the plus rotates 90 degrees. Background, lift, and shadow transitions take 180ms; icons take 200ms. Reduced-motion rules remove these transitions and transforms. General inline actions change color over 120ms without adding a background.

### Cards & File Containers

File cards use 10px corners, 16px padding, and 10px vertical margins. Uploaded cards are filled mist gray without a border; downloadable cards have a thin translucent border. Dark uploads use graphite, and dark downloads use soft charcoal. Card contents form a horizontal row: a fixed 50px square icon tile, flexible file metadata, then compact actions. Icon tiles have 10px corners and 10px padding.

File names use semibold text, with the long leading portion ellipsized and the ending preserved. Status text sits below with 50% opacity and a 4px top / 8px bottom margin. Download actions are blue icon-label pairs with an 18px row gap. These cards remain flat; elevation is reserved for menus and button hover.

### Navigation

A fixed header occupies 10vh. The blue outlined file-lock mark is 22 by 27px and sits 8px before the 26px wordmark. At wide widths, the wordmark and menu align with opposite edges of the central content column.

The menu control is a staggered, right-aligned three-line mark inside a 28 by 18px box. Its lines are 20, 14, and 9px wide and 2px high. It rests at 60% opacity and brightens on hover. Opening rotates the control 180 degrees while folding the lines into an X over 500ms; reduced motion removes the animation.

The dropdown has 10px corners, a 1px translucent border, at least 210px width, and a soft `2px 10px 10px` shadow. Light shadow alpha is `#00000011`; dark is `#00000055`. Blue 11px section captions separate 15px menu entries, each padded 10px vertically and 24px horizontally. Appearance options form a compact Light / Dark / Auto row; the selected option is blue. Their visible keyboard focus uses a 2px blue outline with 2px offset. Account controls appear after authentication and use the same dropdown vocabulary, with a 220px minimum width and a compact identity header.

### Inputs & Forms

Share-key textareas have 10px corners, 16px padding, 16px Inter text, 1.5 line height, a white surface, and a faint 1px border. They grow with content, do not expose a resize handle, and break long key strings. Actions align to the right with a 20px gap.

Nickname inputs use the same radius and border with 12px by 14px padding. Their focus border changes to translucent blue `#1377f955`. The shipped textarea CSS removes the focus outline; a global focus ring for all controls was not established. If extending the design, treat additional accessible focus styling as an intentional improvement, not an extracted property. Dark inputs use soft-charcoal fill, frost text, and translucent white borders.

### Conversation, Drop Dock & Contacts

System messages are unfilled text blocks, indented 50px from their outer column. A circular status avatar sits 48px to the left and 5px above the text edge. The thin tinted border encloses a small 13 by 16px lock. Error and warning messages retain the same shape and change only their icon and border colors.

The bottom action dock is fixed within a full-width page-colored wrapper. Its pale-blue container has 14px corners, 16px by 20px padding, a 16px gap, and 50% viewport width on desktop. A file symbol, semibold title, and button group form one row. The dock is hidden before the applicable file workflow is active; its initial absence is intentional.

Contact chips are compact pills with 5px vertical padding, a 24px circular initial avatar, and a 14px label. Contact managers use a single rounded bordered list, 12px by 14px rows, subtle separators, and small blue edit actions. Destructive actions turn red on hover. Identity graphics stay small and functional.

## 5. Layout Principles

### Grid & Structure

This is a flex-based, single-column conversation with fixed header and bottom dock. There is **no fixed content max-width** in the core stylesheet. Above 1080px, the main container uses `15vh 25% 0` padding, creating a central column equal to 50% of the viewport. Individual message/file wrappers are 50% of that inner column and have a 50px left indent. At a 1440px viewport, the central column is 720px and the initial message width is 360px; retain this unusually narrow desktop conversation when recreating the source faithfully.

The header logo begins at 25% of the viewport and the menu ends at 75%. An opaque fixed top band prevents scrolling content from crossing over the controls. Uploaded items align to the right; system instructions align left. The inner conversation reserves 22vh beneath its content for the fixed dock.

### Whitespace Strategy

The spacing vocabulary is practical rather than a strict 4px/8px token grid: 7, 8, 10, 12, 14, 16, 18, 20, 22, and 24px all occur. Use 16px as a typical card inset, 10px for compact control spacing, and 22px margins around message blocks. Large page spacing depends on the viewport: 15vh top inset, 10vh header height, and a dock bottom margin of 5vh. Preserve the abundant blank canvas around the compact conversation.

### Alignment & Visual Balance

Instructions and filenames are left aligned. The logo balances the small menu control across the header. Blue circular avatars provide a consistent visual anchor beside messages. The dock distributes its visual weight across icon, flexible text, and actions; on phones its title becomes centered. Imagery is limited to small SVG symbols, contact identity marks, and functional file graphics.

### Responsive Behavior & Touch

| Condition | Shipped behavior |
| --- | --- |
| At least 1081px | Central 50% column; header controls align to its edges; opaque 10vh top band |
| At most 1080px | Main padding becomes `15vh 5vw`; message widths become `calc(100% - 70px)`; logo bar spans viewport; menu sits 5vw from right |
| At most 1080px | Dock becomes 80% viewport width; key input, contact picker, and nickname row become 85% of their containing width |
| At most 720px | Dock becomes 90% wide with 16px padding; icon hides; centered title sits above equal-width buttons; title scales from 11–15px |
| Portrait and at most 800px | Notification text shrinks from 20px to 16px; notification inset becomes 20px by 15px |
| Reduced-motion preference | Button/icon/menu transitions are removed |

These are desktop-base styles with narrower-screen overrides. At 390px, the measured side inset is 19.5px and message text width is 281px; the inspected initial page has no horizontal overflow. Buttons wrap when needed. The first-run buttons are about 32px high, so a universal 44px minimum touch target is not part of the shipped design; increasing touch targets would be an extension.

## 6. Design System Notes for Stitch Generation

### Language to Use

“Quiet file-security utility, spacious white canvas, compact conversational instructions, neutral Inter typography, precise electric-blue actions, small outlined lock avatars, softly rounded controls, flat file cards, and a restrained floating menu.” For dark mode add: “Near-black page, charcoal surfaces, frost-white prose, brighter blue links, identical geometry and spacing.”

### Color References

Lead with **Paper White `#ffffff`**, **True Ink `#000000`**, **Electric Blue `#1377f9`**, **Mist Gray `#f4f4f4`**, and **Pale Sky `#eaf2fe`**. For dark generation use **Midnight Ink `#0c0c0e`**, **Soft Charcoal `#161618`**, **Graphite `#1f1f23`**, **Frost White `#f2f2f5`**, and **Bright Link Blue `#4d9bff`**. Use red and amber only for their functional states.

### Component Prompts

1. **First-run screen:** “Create ExcelKey’s quiet file-encryption welcome screen on white. Place a small blue file-lock logo and 26px wordmark in a fixed header with a staggered menu mark opposite. Set instructions as left-aligned 16px Inter conversation blocks with small blue outlined lock avatars and generous surrounding space. Begin with a bold sentence at body size. Below the second message, place compact rounded Unlock and Create buttons with fingerprint and plus symbols; Unlock is electric blue, Create is white with a faint border.”
2. **File-action dock:** “Design a compact bottom file-action dock in pale sky blue with softly rounded 14px corners. Arrange a blue file symbol, semibold task title, and electric-blue primary plus white secondary buttons in one row. On phones, omit the symbol, center the title, and move two equal-width buttons beneath it.”
3. **Dark file conversation:** “Render a file-security conversation on a near-black page with frost-white 16px Inter prose and small blue lock avatars. Show a right-aligned uploaded-file card in graphite and a left-aligned downloadable-file card in soft charcoal, both with 10px corners and compact icon-label actions. Use brighter blue links and a small bordered charcoal menu; preserve the light design’s spacious layout.”

### Incremental Iteration

Establish the page column, 15vh top inset, avatar indentation, and compact typography first. Then match the button shapes, file-card proportions, and restrained color roles. Add the mobile dock rearrangement and header behavior before refining hover motion. Generate dark mode by changing colors while retaining geometry. Keep the initial screen focused on instructions and its two choices; introduce file cards, contact controls, and the dock only in relevant workflow states. Distinguish faithful reproduction from improvements such as broader desktop text columns, larger touch targets, or stronger keyboard focus.
