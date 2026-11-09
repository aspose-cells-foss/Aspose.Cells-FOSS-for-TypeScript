import { Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

const outputPath = sampleOutputPath("document-properties-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Sample";

sheet.Cells["A1"].PutValue("Title");
sheet.Cells["B1"].PutValue("Document Properties Sample");

const properties = workbook.DocumentProperties;
properties.Title = "Annual Sales Report 2024";
properties.Subject = "Financial Performance Analysis";
properties.Author = "Finance Department";
properties.Keywords = "sales, finance, 2024, report";
properties.Comments = "This document contains quarterly sales data and projections";
properties.Category = "Financial Reports";
properties.Company = "Acme Corporation";
properties.Manager = "Jane Smith";

properties.Core.Creator = "Finance Department";
properties.Core.LastModifiedBy = "Finance Department";
properties.Core.Created = new Date();

properties.Extended.Company = "Acme Corporation";
properties.Extended.Manager = "Jane Smith";

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedProperties = loaded.DocumentProperties;

console.log("Saved: " + outputPath);
console.log("Title: " + loadedProperties.Title);
console.log("Subject: " + loadedProperties.Subject);
console.log("Author: " + loadedProperties.Author);
console.log("Keywords: " + loadedProperties.Keywords);
console.log("Comments: " + loadedProperties.Comments);
console.log("Category: " + loadedProperties.Category);
console.log("Company: " + loadedProperties.Company);
console.log("Manager: " + loadedProperties.Manager);
console.log("Core Creator: " + loadedProperties.Core.Creator);
console.log("Core LastModifiedBy: " + loadedProperties.Core.LastModifiedBy);
console.log("Extended Company: " + loadedProperties.Extended.Company);
console.log("Extended Manager: " + loadedProperties.Extended.Manager);
