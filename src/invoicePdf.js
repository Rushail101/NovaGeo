// GST tax invoice / credit note PDF (fields per CGST Rule 46). Built with jsPDF, same as the PO PDFs.
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { STATES, taxBreakup } from "./taxLogic.js";
import { amountInWords, fmtDate } from "./ui.jsx";

const NAVY = [22, 42, 74];
const PALE = [232, 239, 248];
const money = n => (+n || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const TITLES = { invoice: "TAX INVOICE", credit_note: "CREDIT NOTE", debit_note: "DEBIT NOTE" };

function drawCopy(doc, inv, co, copyLabel) {
  const W = doc.internal.pageSize.getWidth();
  const M = 12;
  // Header band
  doc.setFillColor(...NAVY); doc.rect(0, 0, W, 26, "F");
  doc.setTextColor(255); doc.setFont("helvetica", "bold"); doc.setFontSize(15);
  doc.text(co.legal_name || "Nova Geo Resources LLP", M, 11);
  doc.setFont("helvetica", "normal"); doc.setFontSize(7.6);
  doc.text(doc.splitTextToSize(co.address || "", 120), M, 16);
  doc.text(`GSTIN: ${co.gstin || "-"}   PAN: ${co.pan || "-"}${co.phone ? "   Ph: " + co.phone : ""}${co.email ? "   " + co.email : ""}`, M, 23);
  doc.setFont("helvetica", "bold"); doc.setFontSize(13);
  doc.text(TITLES[inv.doc_type] || "TAX INVOICE", W - M, 11, { align: "right" });
  doc.setFontSize(7.5); doc.setFont("helvetica", "normal");
  doc.text(copyLabel, W - M, 16, { align: "right" });
  if (inv.status === "cancelled") {
    doc.setTextColor(200, 30, 30); doc.setFontSize(48); doc.setFont("helvetica", "bold");
    doc.text("CANCELLED", W / 2, 140, { align: "center", angle: 25 });
  }
  doc.setTextColor(20);

  // Meta block
  const pos = `${STATES[inv.place_of_supply] || ""} (${inv.place_of_supply})`;
  const meta = [
    [`${inv.doc_type === "invoice" ? "Invoice" : "Note"} No.`, inv.invoice_no, "Date", fmtDate(inv.invoice_date)],
    ["Place of supply", pos, "Reverse charge", inv.reverse_charge ? "Yes" : "No"],
  ];
  if (inv.against_invoice) meta.push(["Against invoice", inv.against_invoice, "", ""]);
  if (inv.irn) meta.push(["IRN", inv.irn, "Ack No. / Date", `${inv.ack_no || ""} ${inv.ack_date ? fmtDate(inv.ack_date) : ""}`]);
  if (inv.eway_bill_no || inv.vehicle_no) meta.push(["E-way bill", `${inv.eway_bill_no || "-"}${inv.eway_bill_date ? " dt " + fmtDate(inv.eway_bill_date) : ""}`, "Vehicle / LR", `${inv.vehicle_no || "-"}${inv.lr_no ? " / " + inv.lr_no : ""}`]);
  if (inv.transporter) meta.push(["Transporter", inv.transporter, "Due date", inv.due_date ? fmtDate(inv.due_date) : ""]);
  autoTable(doc, {
    startY: 30, margin: { left: M, right: M }, theme: "grid", body: meta,
    styles: { fontSize: 8, cellPadding: 1.6, lineColor: [200, 208, 220] },
    columnStyles: { 0: { fontStyle: "bold", fillColor: PALE, cellWidth: 30 }, 2: { fontStyle: "bold", fillColor: PALE, cellWidth: 30 } },
  });

  // Parties
  const billState = inv.bill_state_code ? `${STATES[inv.bill_state_code] || ""} (${inv.bill_state_code})` : "";
  autoTable(doc, {
    startY: doc.lastAutoTable.finalY + 3, margin: { left: M, right: M }, theme: "grid",
    head: [["Billed to", "Shipped to"]],
    body: [[
      [inv.bill_name, inv.bill_address, billState && `State: ${billState}`, inv.bill_gstin ? `GSTIN: ${inv.bill_gstin}` : "Unregistered"].filter(Boolean).join("\n"),
      [inv.ship_name || inv.bill_name, inv.ship_address || inv.bill_address].filter(Boolean).join("\n"),
    ]],
    headStyles: { fillColor: PALE, textColor: NAVY, fontStyle: "bold", fontSize: 8 },
    styles: { fontSize: 8, cellPadding: 2, lineColor: [200, 208, 220], valign: "top" },
  });

  // Items
  const inter = +inv.igst > 0 || String(inv.place_of_supply) !== String(co.state_code);
  const head = inter
    ? [["#", "Description", "HSN", "Qty", "Unit", "Rate", "Taxable", "IGST %", "IGST", "Total"]]
    : [["#", "Description", "HSN", "Qty", "Unit", "Rate", "Taxable", "GST %", "CGST", "SGST", "Total"]];
  const body = (inv.items || []).map((l, i) => inter
    ? [i + 1, l.description, l.hsn, l.qty, l.unit, money(l.rate), money(l.taxable), `${l.gst_rate}%`, money(l.igst), money(l.total)]
    : [i + 1, l.description, l.hsn, l.qty, l.unit, money(l.rate), money(l.taxable), `${l.gst_rate}%`, money(l.cgst), money(l.sgst), money(l.total)]);
  autoTable(doc, {
    startY: doc.lastAutoTable.finalY + 3, margin: { left: M, right: M }, theme: "grid", head, body,
    headStyles: { fillColor: NAVY, textColor: 255, fontSize: 7.5, halign: "center" },
    styles: { fontSize: 7.8, cellPadding: 1.6, lineColor: [200, 208, 220] },
    columnStyles: { 0: { cellWidth: 7, halign: "center" }, 1: { cellWidth: inter ? 48 : 40 }, 3: { halign: "right" }, 5: { halign: "right" }, 6: { halign: "right" }, 7: { halign: "center" }, 8: { halign: "right" }, 9: { halign: "right" }, 10: { halign: "right" } },
  });

  // Totals + tax breakup
  let y = doc.lastAutoTable.finalY + 3;
  const totals = [["Taxable value", money(inv.taxable_total)]];
  if (inter) totals.push(["IGST", money(inv.igst)]); else totals.push(["CGST", money(inv.cgst)], ["SGST", money(inv.sgst)]);
  if (+inv.round_off) totals.push(["Round off", money(inv.round_off)]);
  totals.push(["Total (Rs.)", money(inv.grand_total)]);
  autoTable(doc, {
    startY: y, margin: { left: W - M - 70, right: M }, theme: "grid", body: totals, tableWidth: 70,
    styles: { fontSize: 8.5, cellPadding: 1.8, lineColor: [200, 208, 220] },
    columnStyles: { 0: { fontStyle: "bold", fillColor: PALE }, 1: { halign: "right" } },
    didParseCell: d => { if (d.row.index === totals.length - 1) { d.cell.styles.fontStyle = "bold"; d.cell.styles.fontSize = 10; } },
  });
  const afterTotals = doc.lastAutoTable.finalY;
  const brk = taxBreakup(inv.items || []);
  autoTable(doc, {
    startY: y, margin: { left: M, right: 84 }, theme: "grid",
    head: [["HSN", "Taxable", "Rate", inter ? "IGST" : "CGST", ...(inter ? [] : ["SGST"])]],
    body: brk.map(b => [b.hsn, money(b.taxable), `${b.rate}%`, money(inter ? b.igst : b.cgst), ...(inter ? [] : [money(b.sgst)])]),
    headStyles: { fillColor: PALE, textColor: NAVY, fontSize: 7.5 },
    styles: { fontSize: 7.5, cellPadding: 1.4, lineColor: [200, 208, 220] },
    columnStyles: { 1: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" } },
  });
  y = Math.max(afterTotals, doc.lastAutoTable.finalY) + 5;

  doc.setFontSize(8.2); doc.setFont("helvetica", "bold");
  doc.text("Amount in words:", M, y);
  doc.setFont("helvetica", "normal");
  doc.text(doc.splitTextToSize(amountInWords(inv.grand_total), W - 2 * M - 28), M + 28, y);
  y += 7;

  // Bank + terms + signature
  const bank = [co.bank_name && `Bank: ${co.bank_name}`, co.bank_account && `A/c No: ${co.bank_account}`, co.bank_ifsc && `IFSC: ${co.bank_ifsc}`, co.bank_branch && `Branch: ${co.bank_branch}`].filter(Boolean);
  const leftW = W - 2 * M - 72;
  doc.setFontSize(7.6);
  if (bank.length) { doc.setFont("helvetica", "bold"); doc.text("Bank details", M, y); doc.setFont("helvetica", "normal"); doc.text(bank, M, y + 4); }
  const termsY = y + (bank.length ? 6 + bank.length * 3.4 : 0);
  if (co.invoice_terms) { doc.setFont("helvetica", "bold"); doc.text("Terms", M, termsY); doc.setFont("helvetica", "normal"); doc.text(doc.splitTextToSize(co.invoice_terms, leftW), M, termsY + 4); }
  if (inv.notes) { const ny = termsY + 6 + doc.splitTextToSize(co.invoice_terms || "", leftW).length * 3.4; doc.text(doc.splitTextToSize(`Note: ${inv.notes}`, leftW), M, ny); }
  doc.setFont("helvetica", "bold"); doc.setFontSize(8.2);
  doc.text(`For ${co.legal_name || "Nova Geo Resources LLP"}`, W - M, y, { align: "right" });
  doc.setFont("helvetica", "normal");
  doc.text(co.signatory || "Authorised Signatory", W - M, y + 22, { align: "right" });
  doc.setDrawColor(...NAVY); doc.line(W - M - 60, y + 18, W - M, y + 18);

  const H = doc.internal.pageSize.getHeight();
  doc.setFontSize(6.8); doc.setTextColor(120);
  doc.text("This is a computer-generated document.", W / 2, H - 6, { align: "center" });
  doc.setTextColor(20);
}

export function invoicePDF(inv, co, { copies = ["ORIGINAL FOR RECIPIENT", "DUPLICATE FOR TRANSPORTER", "TRIPLICATE FOR SUPPLIER"] } = {}) {
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  copies.forEach((c, i) => { if (i) doc.addPage(); drawCopy(doc, inv, co, c); });
  doc.save(`${inv.invoice_no.replace(/[\\/]+/g, "-")}.pdf`);
}
