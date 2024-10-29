function main(workbook: ExcelScript.Workbook) {
  const sheetActualRule = workbook.getWorksheet("ACTUAL RULE");
  const sheetServers = workbook.getWorksheet("Servers") || workbook.addWorksheet("Servers");
  const sheetServices = workbook.getWorksheet("Services") || workbook.addWorksheet("Services");
  const sheetConnections = workbook.getWorksheet("Connections") || workbook.addWorksheet("Connections");

  // Pobranie danych z ACTUAL RULE
  const actualRuleData = sheetActualRule.getUsedRange().getValues();

  // Zbiór unikalnych adresów IP i portów
  const uniqueSources = new Set<string>();
  const uniqueDestinations = new Set<string>();
  const uniqueServices = new Set<string>();
  const lineInfoMap = new Map<string, string[]>(); // Map do przechowywania linii dla adresów IP

  // Przetwarzanie danych ACTUAL RULE
  for (let i = 1; i < actualRuleData.length; i++) {
    const sourceIPs = (actualRuleData[i][0] as string).split("\n");
    const destinationIPs = (actualRuleData[i][1] as string).split("\n");
    const services = (actualRuleData[i][2] as string).split("\n");
    const line = actualRuleData[i][3] as string;

    // Dodaj źródła i cele do unikalnych zbiorów oraz linii
    sourceIPs.forEach(ip => {
      uniqueSources.add(ip);
      if (!lineInfoMap.has(ip)) lineInfoMap.set(ip, []);
      lineInfoMap.get(ip)?.push(line);
    });

    destinationIPs.forEach(ip => {
      uniqueDestinations.add(ip);
      if (!lineInfoMap.has(ip)) lineInfoMap.set(ip, []);
      lineInfoMap.get(ip)?.push(line);
    });

    services.forEach(service => uniqueServices.add(service));
  }

  // Zapisz unikalne adresy w arkuszu Servers
  const serversData = Array.from(lineInfoMap.entries()).flatMap(([ip, lines]) => {
    const details = ip.startsWith("net_") ? ip.slice(4).replace(/_/g, "/") : ip; // Zamiana net_10.99.0.0_24 na 10.99.0.0/24
    return lines.map((line, index) => [index + 1, ip, details, line]); // Dodaj linię dla każdego adresu
  });

  sheetServers.getRange("A1:D1").setValues([["id", "name", "details", "line"]]);
  sheetServers.getRange(`A2:D${serversData.length + 1}`).setValues(serversData);

  // Zapisz unikalne porty w arkuszu Services
  const lastServiceID = getLastID(sheetServices);
  const servicesData: (string | number)[][] = [];
  let nextServiceID = lastServiceID + 1;

  uniqueServices.forEach((service) => {
    servicesData.push([nextServiceID++, service]);
  });

  if (servicesData.length > 0) {
    if (sheetServices.getUsedRange()?.getRowCount() === 0) {
      sheetServices.getRange("A1:B1").setValues([["ID", "name"]]);
    }
    sheetServices.getRange(`A${lastServiceID + 2}:B${lastServiceID + servicesData.length + 1}`).setValues(servicesData);
  }

  // Zasil arkusz Connections
  const connectionsData = actualRuleData.slice(1).map((row, index) => [
    index + 1,
    row[0],  // Source
    row[1],  // Destination
    row[2],  // Service
    row[3]   // Line
  ]);

  sheetConnections.getRange("A1:E1").setValues([["id", "source", "destination", "service", "line"]]);
  sheetConnections.getRange(`A2:E${connectionsData.length + 1}`).setValues(connectionsData);
}

// Funkcja sprawdzająca, czy ciąg znaków to adres IP
function isIPAddress(ip: string): boolean {
  const ipPattern = /^(\d{1,3}\.){3}\d{1,3}$/;
  return ipPattern.test(ip);
}

// Funkcja do pobierania ostatniego ID w arkuszu
function getLastID(sheet: ExcelScript.Worksheet): number {
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return 0;
  const values = usedRange.getValues();
  return values.length > 1 ? Math.max(...values.slice(1).map(row => row[0] as number)) : 0;
}
