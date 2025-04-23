function excelDateToJS(serial) {
    if (!serial || typeof serial !== "number") return serial;
    const utc_days = Math.floor(serial - 25569);
    const date_info = new Date(utc_days * 86400 * 1000);
    return date_info.toISOString().split("T")[0];
  }
  
  async function generateInvoices() {
    const file = document.getElementById('excelFile').files[0];
    if (!file) return alert("Upload an Excel file!");
  
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
  
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const invoice = document.createElement('div');
      invoice.className = 'invoice';
  
      const taxable = parseFloat(row["Taxable Value"] || 0);
      const igst = parseFloat(row["IGST"] || 0);
      const cgst = parseFloat(row["CGST"] || 0);
      const sgst = parseFloat(row["SGST"] || 0);
  
      const rowsHtml = [];
      rowsHtml.push(`<tr>
        <td>1</td>
        <td>${row["Original Supplier GSTIN"]}</td>
        <td>${row["Original supplier Invoice No"]}</td>
        <td>${row["Rate"]}</td>
        <td>${taxable.toFixed(2)}</td>
        <td>${igst.toFixed(2)}</td>
        <td>${cgst.toFixed(2)}</td>
        <td>${sgst.toFixed(2)}</td>
      </tr>`);
  
      for (let j = 1; j < 4; j++) {
        rowsHtml.push(`<tr>
          <td>${j + 1}</td>
          <td colspan="7">&nbsp;</td>
        </tr>`);
      }
  
      invoice.innerHTML = `
        <h2>TAX INVOICE UNDER RULE 54 FOR DISTRIBUTION OF INPUT TAX CREDIT</h2>
        <h3>UNDER RULE 39 OF CGST RULES, 2017</h3>
  
        <div class="row">
          <div class="column">
            <p><strong>Document number: -</strong> ${row["Document Number"] || ""}</p>
            <p><strong>Document Date: -</strong> ${excelDateToJS(row["Document Date"] || "")}</p>
            <p><strong>Details of Normal Registration: -</strong></p>
            <p><strong>Name:</strong> ${row["Normal Registration Name"]}</p>
            <p><strong>Address:</strong> ${row["Normal Registration Address"]}</p>
            <p><strong>Pin code:</strong> ${row["Normal Registration Pincode"]}</p>
            <p><strong>State Name:</strong> ${row["Normal Registration State"]}</p>
            <p><strong>State code:</strong> ${row["Normal Registration State Code"]}</p>
            <p><strong>GSTIN:</strong> ${row["Normal GSTIN"]}</p>
          </div>
  
          <div class="column">
            <br><br><br>
            <p><strong>Details of ISD Recipient: -</strong></p>
            <p><strong>Name:</strong> ${row["ISD Recipient Name"]}</p>
            <p><strong>Address:</strong> ${row["ISD Recipient Address"]}</p>
            <p><strong>Pin code:</strong> ${row["ISD Recipient Pincode"]}</p>
            <p><strong>State Name:</strong> ${row["ISD Recipient State"]}</p>
            <p><strong>State code:</strong> ${row["ISD Recipient State Code"]}</p>
            <p><strong>GSTIN:</strong> ${row["ISD GSTIN"]}</p>
          </div>
        </div>
  
        <table>
          <thead>
            <tr>
              <th>Sr. No.</th>
              <th>Original Supplier GSTIN</th>
              <th>Original supplier Invoice No</th>
              <th>Rate</th>
              <th>Taxable value</th>
              <th>IGST</th>
              <th>CGST</th>
              <th>SGST</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml.join("\n")}
            <tr>
              <td colspan="4"><strong>Total Credit to be transferred</strong></td>
              <td><strong>${taxable.toFixed(2)}</strong></td>
              <td><strong>${igst.toFixed(2)}</strong></td>
              <td><strong>${cgst.toFixed(2)}</strong></td>
              <td><strong>${sgst.toFixed(2)}</strong></td>
            </tr>
          </tbody>
        </table>
  
        <div class="footer">
          <br><br><br><p><strong>For ${row["Normal Registration Name"]}</strong></p>
          <br>
          <br>
          <br>
          <br><p>(Authorised Signatory)</p>
          <br><p><strong>${row["Reg Office"]}</strong></p>
          <p><strong>E-mail:</strong> ${row["E Mail"]}, <strong>Website:</strong> ${row["Website"]}</p>
          <p><strong>CIN:</strong> ${row["CIN"]}</p>
        </div>
      `;
  
      document.body.appendChild(invoice);
  
      await html2pdf().from(invoice).set({
        filename: `Invoice_${i + 1}_${row["Document Number"]}.pdf`,
        jsPDF: { format: 'a4' },
        html2canvas: { scale: 2 },
        margin: 0
      }).save();
  
      document.body.removeChild(invoice);
    }
  }
  