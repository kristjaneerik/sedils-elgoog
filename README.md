# sedils-elgoog aka `"google slides"[::-1]`

Add reversed slide numbers to Google Slides.

# Usage

```bash
python -m sedils_elgoog presentation_id
```

Find the `presentation_id` from the URL of a presentation.
For example, for https://docs.google.com/presentation/d/1jvtFxrzqQKiSCqeAm5Nggg-JAReXV5Hf4hjos1yNy9A/edit,
the `presentation_id` is `1jvtFxrzqQKiSCqeAm5Nggg-JAReXV5Hf4hjos1yNy9A`, so you would use:

```bash
python -m sedils_elgoog 1jvtFxrzqQKiSCqeAm5Nggg-JAReXV5Hf4hjos1yNy9A
```

## Flags

- `--last-is-zero`: the last slide is number 0, rather than number 1 (default)
- `--remove-all`: remove all previously added slide numbers
- `--color=your_color_choice`: one of https://developers.google.com/slides/reference/rest/v1/presentations.pages/other#Page.ThemeColorType
  Applies to all numbers
- `--fmt="your format string"`: a Python format string, e.g. `{n}` (the default) or `#{n}/{total}`.
  You have access to the variables `n` (number of the current slide)
  and `total` (total number of slides)

## Control in speaker notes

You can control aspects of the numbering of slides by adding certain keywords in the speaker notes
section of a slide.

- `@no-number`: don't show the slide number on this slide. It is still counted
- `@skip-count@`: do not count this slide in the total and don't show a slide number
- `@supplement-start@`: do not count this slide or any slide after it.
   Use it to e.g. not have your supplemental slides be counted toward the total

# TODOs

- [x] requirements.txt
- [ ] make it pip-installable
- [ ] documentation on installation
- [ ] documentation on usage
- [ ] Google API error handling / retry logic
- [ ] per-slide color
