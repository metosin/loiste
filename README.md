# Loiste

An Excel library for Clojure.

---

Verb<br/>
**Loistaa (Excel)**

1. (intransitive) To shine, glow, glare, sparkle (to emit light).
2. (intransitive) To excel, shine (to distinguish oneself).

Nouns: Loiste

---

## Usage

...

## Ideas

- Streaming worksbooks (`SXSSFWorkbook`)
- New workbooks & from template
- Should support lazy seq as data
- Input as Hiccup style data

```clj
[:workbook
 {:styles {:date {:data-format (loiste/date-format (loiste/locale :fi) "dd.MM.yyyy HH:mm")}}}
 [:sheet {:name "Sheet a"}
  [:columns
   [:column {:style :date}]
   [:column]]
  [:rows
   [:row
    [:cell #DateTime "2015-11-20T09:43:00Z"]
    [:cell "foobar"]]]]]

;; Seq/list splicing
[:rows (map (fn [i] [:row [:cell i]]) (range 100))]
```
