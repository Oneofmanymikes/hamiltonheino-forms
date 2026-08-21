# Heino Doessing — form embed pack

Copy-paste code for the campaign's five public forms. Give a web developer
whichever blocks they need. Kept in the repo so it is version-controlled and
cannot go missing.

All five are served by ONE Google Apps Script web app. The base URL is identical
for every form — only the `?form=` parameter changes:

```
https://script.google.com/macros/s/AKfycbxDI4llh5qvVA-fSw3UuogwNLhQuv9kChghwmV0-h2bdsnZZug6JW5QILRZF09hdcAa/exec?form=<vote|sign|donate|volunteer|shift>
```

| # | Form | Height | `?form=` |
|---|---|---|---|
| 1 | Commit to Vote | 560px | `vote` |
| 2 | Request a Lawn Sign | 900px | `sign` |
| 3 | Commit to Donate (pledge only) | 680px | `donate` |
| 4 | Volunteer Sign-Up | 760px | `volunteer` |
| 5 | Sign Up to Canvass (live shift booking) | 1000px | `shift` |

**#4 and #5 are not duplicates.** #4 is a general "I'd like to help". #5 books a
named person into a dated canvass shift with live capacity.

---

## The embed code

Replace `FORM` with the `?form=` value and `HEIGHT` with the height above.

```html
<iframe
  src="https://script.google.com/macros/s/AKfycbxDI4llh5qvVA-fSw3UuogwNLhQuv9kChghwmV0-h2bdsnZZug6JW5QILRZF09hdcAa/exec?form=FORM"
  title="Heino Doessing campaign form"
  style="width:100%;max-width:560px;height:HEIGHTpx;border:0;"
  loading="lazy"></iframe>
```

In WordPress use a **Custom HTML** block — a Paragraph block escapes the tags
and shows the code as text. Squarespace: a Code block. Wix: Embed → Embed HTML.

## Direct links (no embedding)

For emails, texts and social posts, where iframes do not render:

```
Commit to Vote    https://voteheino.ca/getinvolved/commit-to-vote/
Lawn Sign         https://voteheino.ca/getinvolved/lawn-sign/
Commit to Donate  https://voteheino.ca/getinvolved/donate/
Volunteer         https://voteheino.ca/getinvolved/volunteer/
Canvass shifts    https://voteheino.ca/getinvolved/shift/
All of them       https://voteheino.ca/getinvolved/
```

---

## Canvass shift times (form 5)

Each block runs two hours.

| Day | Start times |
|---|---|
| Monday–Friday | 1:00 PM, 5:00 PM |
| Saturday | 10:00 AM, 1:00 PM, 3:00 PM, 5:00 PM |
| Sunday | 12:00 noon, 3:00 PM, 5:00 PM |

## Things that look like bugs but are not

- **Heights are fixed.** A cross-origin iframe cannot report its own height, so
  these cannot auto-size. Adjust the `height` for your theme; that is the only
  number to touch. Form 5 is the tallest and will scroll internally — expected.
- **Form 5 recognises returning volunteers.** Type a known phone number and the
  name/email fields disappear, replaced by "Welcome back, <name>". Intentional.
- **Full shifts vanish from the list**, and shifts with 3 or fewer places left
  show a "3 left" badge. Two people can legitimately see different options.
- **The donate form is a PLEDGE form.** No card details are collected. Do not
  relabel it in a way that implies a visitor is being charged; real donations go
  through `ontarioliberal.ca/donate/?pla=37`.

## Notes

- Works on any domain — no CORS setup, no API keys, no accounts.
- **Never edit the `src` beyond the `?form=` value.** The long string is the
  deployment ID. If the campaign ever redeploys to a NEW deployment instead of
  updating the existing one, that ID changes and every embed breaks at once.
- Nothing is stored on the host site: submissions post straight to Google, so
  there is no personal data on your server.
- A hidden honeypot field silently discards bot submissions, and shift capacity
  is enforced server-side under a lock so a shift cannot be over-booked.
