const readXlsxFile = require("read-excel-file/node");

const generateCode = (name) => {
  var newCode = "";

  if (name.includes('"')) {
    newCode = name.split('"').pop().split('"')[0];

    newCode = newCode.replace("s", "-");
  } else {
    var nameParts = name.split(/[\s-]/gm);

    if (nameParts.length === 2) {
      newCode = nameParts[0] + "-" + nameParts[1];
    } else if (nameParts.length > 1 && nameParts.length < 6) {
      nameParts.forEach((part) => {
        newCode += part.slice(0, 1);
      });
    } else if (name.length < 12) {
      newCode = name;
    } else {
      newCode = name.slice(0, 3);
    }
  }

  newCode = newCode.toUpperCase();
  return newCode;
};

const contractorsParser = async () => {
  let contractors = [];
  await readXlsxFile("Adresy.xlsx").then((rows) => {
    let contractor = {};
    rows.forEach((row, index) => {
      if (index > 0) {
        row.map((column, index) => {
          switch (index) {
            case 0:
              contractor = {
                ...contractor,
                name: column.replaceAll(/\n/gm, " "),
                code: generateCode(column),
              };
              break;
            case 1:
              let country = "PL";
              const dataToSplit = column.replace("NIP:", "NIP");
              let address;
              let vatId;
              if (dataToSplit.includes("NIP")) {
                address = dataToSplit.split("NIP")[0];
                vatId = dataToSplit.split("NIP")[1];
                vatId = vatId?.replace("NIP", "").replaceAll(" ", "");
              } else if (dataToSplit.includes("(TAX number)")) {
                address = dataToSplit.split("(TAX number)")[0];
                vatId = dataToSplit.split("(TAX number)")[1];
                vatId = vatId.replace("(TAX number)", "").replaceAll(" ", "");
                country = vatId.slice(0, 2);
              } else {
                address = dataToSplit;
              }
              const zipCodeRegex = /[0-9]{2}-[0-9]{3}/gm;
              commaIndex = address?.search(zipCodeRegex) + 6;
              address =
                address?.slice(0, commaIndex) +
                "," +
                address?.slice(commaIndex);
              contractor = {
                ...contractor,
                address,
                country,
                vat_id: vatId,
              };
              break;
            case 2:
              const newLine = /\r?\n/;
              emails = column
                .replaceAll(" ", "")
                .replaceAll(";", ",")
                .split(",");
              if (emails.length === 1) emails = emails[0].split(newLine);
              const emailRegex =
                /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
              emails = emails.filter((email) => emailRegex.test(email));
              contractor = {
                ...contractor,
                email: emails,
              };
              break;
          }
        });
        contractors.push(contractor);
      }
    });
  });
  return contractors;
};

const resp = async () => {
  dane = await contractorsParser();
  console.log(dane);
};

resp();
