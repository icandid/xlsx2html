const XLSX = require("xlsx")
const cheerio = require("cheerio")

const xlsx_file = "./excel_style.xlsx"

const workbook = XLSX.readFile(xlsx_file, {
  cellStyles: true,
})
const html = []

workbook.Strings.forEach(({r}) => {
  if (!r) return

  const $ = cheerio.load(r, null, false)
  const result = []

  function createSpan(text, style) {
    const s = style ? ` style="${style}"` : ""
    return `<span${s}>${text}</span>`
  }

  function getStyle(styleElm) {
    const color = $("color", styleElm).attr("rgb")
    const fontFamily = $("rFont", styleElm).attr("val")
    const fontSize = $("sz", styleElm).attr("val")
    const style = []

    style.push(`color: #${color ? color.slice(2) : "000000"};`)
    if (fontFamily) {
      style.push(`font-family: ${fontFamily};`)
    }
    if (fontSize) {
      style.push(`font-size: ${fontSize}px;`)
    }
    if ($("b", styleElm).length) {
      style.push(`font-weight: bold;`)
    }
    if ($("strike", styleElm).length) {
      style.push("text-decoration: line-through;")
    }
    return style.join(" ")
  }

  $("r").each((i, elm) => {
    const styleElm = $("rPr", elm)
    const style = styleElm ? getStyle(styleElm) : ""
    const text = $("t", elm).text()
    const content = text == "\n" ? "<br />" : createSpan(text, style)
    result.push(content)
  })

  html.push(result.join(""))
})

console.log(html)
