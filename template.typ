// template.typ — Hymnal Edition 2 Typst pilot, presentation layer.
//
// Static / hand-authored. Reads no editorial data itself; main.typ passes
// in hymn dictionaries loaded from hymnal_data.json. Do not put hymn text
// or metadata into this file — that belongs in the workbook.
//
// SCHEMA CONTRACT with transform.py's JSON — if you rename a field in
// transform.py's build_hymn_data(), it must change here too:
//   printed_no, id, title, category, flags (unused here),
//   words_line (string | none), translator_line (string | none),
//   tune_lines (array of string), alt_tunes_line (string | none),
//   sections (array of {type,label,text,stanza_no?}),
//   force_break_before (bool, optional, added by build.py)
// Top-level keys: hymns, indexes, title_page, category_breaks.
// indexes: category (dict), title (array), person_source (dict),
//   tune (array of {name, composer, hymn_nos}).
//
// LOAD-BEARING TAGS: <hymn-debug>, <hymn-measured> and <category-debug> are
// read by pagination.py. Do not remove or rename them without updating it.
// <section-debug> is diagnostic only.
 
// ============================================================
// TUNABLE VALUES — all working, none final. Change values here;
// exceptions: hymn-gap is currently unused, and hymn-block hard-codes the A5 height (595.28pt).
// ============================================================
 
#let page-margin-top-bottom = 0.5cm
#let page-margin-sides = 0.75cm
#let column-gutter = 0.25cm

#let heading-size = 10pt
#let heading-number-gutter = 4pt // space between "No." and title in the heading grid
#let body-size = 9pt
#let attribution-size = 6pt
 
#let body-leading = 0.5em  // Gap between authored lines within one stanza/section. Current working value.
#let authored-line-spacing = 5pt
 
#let stanza-number-col = 12pt
#let stanza-wrap-indent = 8pt   // wrap-indent used inside stanzas
#let label-wrap-indent = 10pt  // wrap-indent inside Chorus/Bridge/etc. (differs from stanza-wrap-indent; undecided whether they should match)
 
#let section-gap = 12pt      // between stanzas / chorus / bridge / etc.
#let attribution-gap = 16pt  // before attribution — MUST stay larger than
                              // section-gap per the typography spec
#let hymn-gap = 12pt         // NOT USED at present; spacing between hymns comes from other block spacing

#let title-gap = 10pt   // gap between title and subtitle

#let index-category-side-margin = 2.5cm
#let index-category-title-gap = 24pt
 
// ============================================================
// PAGE SETUP
// ============================================================
 
#let setup-page(body) = {
  set page(
    paper: "a5",
    margin: (
      top: page-margin-top-bottom, bottom: page-margin-top-bottom,
      left: page-margin-sides, right: page-margin-sides,
    ),
    columns: 2,
  )
  set columns(gutter: column-gutter)
  set text(font: "PT Sans", size: body-size, hyphenate: false)
  body
}
 
#let setup-index-page(body) = {
  set page(  
    paper: "a5",
    margin: (
      top: page-margin-top-bottom, bottom: page-margin-top-bottom,
      left: index-category-side-margin, right: index-category-side-margin,
    ),
    columns: 1,
  )
  set text(font: "PT Sans", size: body-size, hyphenate: false)
 
  body
 
}
 
#let setup-title-page(body) = {
  set page(
    paper: "a5",
    margin: (
      top: page-margin-top-bottom, bottom: page-margin-top-bottom,
      left: page-margin-sides, right: page-margin-sides,
    ),
    columns: 1,
  )
  set text(font: "PT Sans", hyphenate: false)
  set align(center + horizon)
  body
}

#let title-page(info) = setup-title-page[
  #set par(spacing: 0pt, leading: 0.4em)
  #text(size: 24pt, weight: "bold")[#upper(info.title)]
  #v(title-gap)
  #text(size: 14pt, style: "italic")[#info.subtitle]
]
 
// ============================================================
// LINE RENDERING
// Splits text on "\n" (each authored line, per transform.py's
// normalize_lyric_text) into its own paragraph. hanging-indent only
// engages when Typst physically WRAPS one of these single-line
// paragraphs — never on a transition to a new authored line, because
// each authored line already starts its own paragraph via parbreak().
// justify is off: short, deliberately-broken lyric lines look wrong
// stretched to fill the column width.
// ============================================================
 
#let render-lines(txt, wrap-indent: stanza-wrap-indent) = {
  let lines = txt.split("\n")
  for line in lines {
    set par(
      hanging-indent: wrap-indent,
      justify: false,
      leading: body-leading,
      spacing: authored-line-spacing,
    )
    line
    parbreak()
  }
}
 
// ============================================================
// HYMN HEADING
// Grid-based: column 1 auto-sizes to the number's own width, column 2
// (1fr) holds the title. A wrapped title's continuation line lands at
// the start of column 2 — i.e. under the title text, never under the
// number — without needing a manually-tuned indent value at all.
// ============================================================
 
#let hymn-heading(hymn) = {
  context [
    #metadata((
      hymn: hymn.printed_no,
      page: here().position().page,
      x: here().position().x,
      y: here().position().y,
    )) <hymn-debug>
  ]
 
  block(width: 100%, below: section-gap)[
    #set text(
      font: "PT Sans",
      weight: "bold",
      size: heading-size,
      hyphenate: false,
    )
    #grid(
      columns: (auto, 1fr),
      gutter: heading-number-gutter,
      [#hymn.printed_no.],
      [#upper(hymn.title)],
    )
  ]
}
 
// ============================================================
// SECTIONS
// Stanza: numbered, plain (not bold) number, in its own grid column.
// Labelled sections (Chorus / Intro / Outro / Bridge / Final Chorus):
// whole section italicized, label on its own line, sharing the stanza
// text column (empty first grid cell — not flush-left like a number).
// ============================================================
 
#let section-heading-text(section) = {
  if section.label != none {
    section.label
  } else if section.type == "final chorus" {
    "Final Chorus:"
  } else if section.type == "chorus" {
    "Chorus:"
  } else {
    // intro, outro, bridge
    upper(section.type.slice(0, 1)) + section.type.slice(1) + ":"
  }
}
 
#let section-block(section, number-col-width: stanza-number-col) = {
  block(
    above: section-gap,
    below: section-gap,
    breakable: false,
  )[
    #if section.type == "stanza" [
      #grid(
        columns: (number-col-width, 1fr),
        gutter: 0pt,
        text(weight: "regular")[#str(section.stanza_no).],
        render-lines(section.text, wrap-indent: stanza-wrap-indent),
      )
    ] else [
      #grid(
        columns: (number-col-width, 1fr),
        gutter: 0pt,
        [],
        emph[
          #block(below: 5pt)[#section-heading-text(section)]
          #render-lines(section.text, wrap-indent: label-wrap-indent)
        ],
      )
    ]
  ]
}
 
// ============================================================
// ATTRIBUTION (after lyrics, small italic)
// Uses ONLY the fields transform.py actually emits. words_line already
// encodes the "Words and Tune by X" consolidation when author and
// composer match — do not re-derive that check here.
// ============================================================
 
#let attribution-block(hymn) = {
  let lines = ()
  if hymn.words_line != none { lines.push(hymn.words_line) }
  if hymn.translator_line != none { lines.push(hymn.translator_line) }
  for t in hymn.tune_lines { lines.push(t) }
  if hymn.alt_tunes_line != none { lines.push(hymn.alt_tunes_line) }
 
  if lines.len() > 0 {
    set text(size: attribution-size, style: "italic")
    set align(right)
 
    block(
      width: 90%,
      above: attribution-gap,
      below: 4pt,
      breakable: false,
    )[
      #lines.join([ \ ])
    ]
  }
}
 
// ============================================================
// INDEXES
// ============================================================
//
// Python builds the index data. These functions only control
// presentation.
//
// Index 1: Category — front matter
// Index 2: Title — back matter
// Index 3: Author / Source / Translator — back matter
// Index 4: Tune — back matter
//
// The indexes deliberately receive already-resolved data. No
// contributor/tune attribution logic belongs here.
// ============================================================
 
#let index-heading(title) = {
  place(
    top + center,
    scope: "parent",
    float: true,
  )[
    #set text(
      font: "PT Sans",
      weight: "bold",
      size: 18pt,
    )
    #title
  ]
}
 
#let index-row(
  left-content,
  numbers,
  size: 7pt,
  indent: 8pt,
) = {
  block(
    below: 4pt,
    breakable: false,
  )[
    #set text(size: size)
    #par(hanging-indent: indent, justify: false, leading: 0.5em)[
      #left-content #box(width: 1fr, repeat[.#h(2pt)]) #numbers
    ]
  ]
}
 
#let category-index(index, breaks: ()) = {
  let keep = 2   // entries kept with the heading, and minimum carried over
  let category-gap = index-category-title-gap   // space between end of one category and the next heading

  let category-heading(category) = block(
    below: 18pt,
  )[
    #set text(weight: "bold", size: 14pt)
    #category
  ]

  let category-entry(entry) = block(below: 5pt)[
    #set text(size: 10pt)
    #entry.title #box(width: 1fr, repeat[.#h(2pt)]) #entry.no
  ]

  // Position marker read by pagination.py. Placed inside the first block
  // of each category so it reports the heading's real position, after the gap.
  let category-marker(category) = context [
    #metadata((
      category: category,
      page: here().position().page,
      y: here().position().y,
    )) <category-debug>
  ]

  align(center)[
    #index-heading("Category")

    #for (category, entries) in index {
      if category in breaks {
        pagebreak(weak: true)
      }

      let n = entries.len()

      if n <= 2 * keep {
        block(breakable: false, above: index-category-title-gap)[
          #category-marker(category)
          #category-heading(category)
          #for e in entries [#category-entry(e)]
        ]
      } else {
        block(breakable: false, above: index-category-title-gap, below: 5pt)[
          #category-marker(category)
          #category-heading(category)
          #for e in entries.slice(0, keep) [#category-entry(e)]
        ]
        for e in entries.slice(keep, n - keep) {
          category-entry(e)
        }
        block(breakable: false)[
          #for e in entries.slice(n - keep) [#category-entry(e)]
        ]
      }
    }
  ]
}

#let title-index(index) = {
  index-heading("Title Index")

  for entry in index {
    index-row(entry.title, str(entry.no))
  }
}
 
#let person-source-index(index) = {
  index-heading("Index by Author, Translator, and Source")
 
  for (name, hymn_nos) in index {
    index-row(name, hymn_nos.map(str).join(", "))
  }
}
 
#let tune-index(index) = {
  index-heading("Index by Tunes (with Composer or Source)")
 
  for entry in index {
    let left = emph[#quote(block: false)[#entry.name]] + if entry.composer != "" [
      #h(0.1em)#sym.hyph#h(0.2em)#entry.composer
    ] else []
    index-row(left, entry.hymn_nos.map(str).join(", "))
  }
}
 
// ============================================================
// FULL HYMN
// ============================================================
 
#let hymn-block(hymn) = {
  let hymn-content = [
    #hymn-heading(hymn)
 
    #for (idx, section) in hymn.sections.enumerate() {
      context [
        #metadata((
          hymn: hymn.printed_no,
          section-index: idx,
          section-type: section.type,
          page: here().position().page,
          x: here().position().x,
          y: here().position().y,
        )) <section-debug>
      ]
      if idx == hymn.sections.len() - 1 {
        block(breakable: false)[
          #section-block(section)
          #attribution-block(hymn)
        ]
      } else {
        section-block(section)
      }
    }
  ]
 
  if hymn.at("force_break_before", default: false) {
    colbreak(weak: true)
  }
 
  layout(size => {
    let measured = measure(
      width: size.width,
      hymn-content,
    )
 
    let column-height = 595.28pt - 2 * page-margin-top-bottom
 
    let body = if measured.height <= column-height {
      block(breakable: false)[
        #hymn-content
      ]
    } else {
      hymn-content
    }
 
    [
      #metadata((hymn: hymn.printed_no, measured-height: measured.height)) <hymn-measured>
      #body
    ]
  })
}
