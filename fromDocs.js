const PizZip = require("pizzip")
const Docxtemplater = require("docxtemplater")
const express = require("express")
const fs = require("fs")
const path = require("path")
const app = express()
app.use(express.urlencoded())
const port = 3000

let resourceDesignation = ""
let resourceName = ""

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"))
})

function editWord(jsonInput) {
  console.log(jsonInput["Start Date"])
  productName = jsonInput["Product Name"]
  sequenceNumber = jsonInput["Sequence Number"]
  customerName = jsonInput["Customer Name"]
  companyName = jsonInput["Company Name"]
  // startEndDates = jsonInput["Start Date - End Date"]
  startDate = jsonInput["Start Date"]
  endDate = jsonInput["End Date"]
  country = jsonInput["Country"]
  customerLocation = jsonInput["Customer Location"]
  supplierContactName = jsonInput["Supplier Contact Name"]
  supplierTitle = jsonInput["Supplier Title"]
  supplierPhoneNumber = jsonInput["Supplier Phone Number"]
  supplierEmail = jsonInput["Supplier Email Address"]
  supplierName = jsonInput["Supplier Name"]

  console.log(startDate, " ", endDate)

  let i = 1
  let resource = ""
  let resourceName = ""
  // resourceDesignation = jsonInput["resource1"]
  // resourceName = jsonInput["resourceName1"]

  resource = jsonInput["resource" + i] //hello(world)
  resourceName = " (" + jsonInput["resourceName" + i] + ")\n"
  i++

  while (jsonInput["resource" + (i + 1)] != undefined) {
    resourceName +=
      jsonInput["resource" + i] + " (" + jsonInput["resourceName" + i] + ")\n"
    console.log(resource)
    i++
  }

  resourceName +=
    jsonInput["resource" + i] + " (" + jsonInput["resourceName" + i] + ")"

  const content = fs.readFileSync(
    path.resolve(__dirname, "SOW Template.docx"),
    "binary"
  )

  const zip = new PizZip(content)

  const doc = new Docxtemplater(zip, {
    delimiters: { start: "<", end: ">" },
    paragraphLoop: true,
    linebreaks: true,
  })

  doc.render({
    "Product Name": productName,
    "Sequence Number": sequenceNumber,
    "Customer Name": customerName,
    "Company Name": companyName,
    "Start Date - End Date": `${startDate} - ${endDate}`,
    Country: country,
    "Customer Location": customerLocation,
    "Supplier Contact Name": supplierContactName,
    "Supplier Title": supplierTitle,
    "Supplier Phone Number": supplierPhoneNumber,
    "Supplier Email Address": supplierEmail,
    "Supplier Name": supplierName,
    "Resource Designation": resource,
    "Resource Name": resourceName,
  })

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  })

  fs.writeFileSync(path.resolve(__dirname, "New SOW Template.docx"), buf)
}

app.post("/process", (req, res) => {
  console.log(req.body)
  editWord(req.body)
  res.redirect("/success")
})

app.get("/DocFile.docx", (req, res) => {
  res.sendFile(path.join(__dirname, "output.docx"))
})

app.listen(3000, (req, res) => {
  console.log("server is up")
})
