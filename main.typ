// main.typ — Hymnal Edition 2 Typst pilot entry point.
//
// Loads the data Python produced and renders it via template.typ.
// This file controls document structure/order only.
// Presentation details belong in template.typ.

#import "template.typ": (
  setup-page,
  setup-index-page,
  title-page,
  hymn-block,
  category-index,
  title-index,
  person-source-index,
  tune-index,
)

#let data = json("hymnal_data.json")

// ============================================================
// FRONT MATTER
// ============================================================

#setup-index-page[
  #category-index(data.indexes.category)
]

// ============================================================
// TITLE PAGE
// ============================================================
 
#title-page(data.title_page)

// ============================================================
// HYMN BODY
// ============================================================

#setup-page[
  #for hymn in data.hymns [
    #hymn-block(hymn)
  ]
]

// ============================================================
// BACK MATTER — INDEXES
// ============================================================

#pagebreak()

#setup-page[
  #title-index(data.indexes.title)
]

#pagebreak()

#setup-page[
  #person-source-index(data.indexes.person_source)
]

#pagebreak()

#setup-page[
  #tune-index(data.indexes.tune)
]
