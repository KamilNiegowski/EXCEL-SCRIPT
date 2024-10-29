function main(workbook: ExcelScript.Workbook) {
  const sheetActualRule = workbook.getWorksheet("ACTUAL RULE");
  const sheetServers = workbook.getWorksheet("Servers") || workbook.addWorksheet("Servers");
  const sheetServices = workbook.getWorksheet("Services") || workbook.addWorksheet("Services");
  const sheetConnections = workbook.getWorksheet("Connections") || workbook.addWorksheet("Connections");

  // Pobranie danych z ACTUAL RULE
  const actualRuleData = sheetActualRule.getUsedRange().getValues();

  // Zbiór unikalnych adresów IP i portów
  const uniqueServers = new Map<string, Set<string>>(); // key: IP/name, value: Set of lines
  const uniqueServices = new Set<string>();

  // Przetwarzanie danych ACTUAL RULE
  for (let i = 1; i < actualRuleData.length; i++) {
    const sourceIPs = (actualRuleData[i][0] as string).split("\n");
    const destinationIPs = (actualRuleData[i][1] as string).split("\n");
    const services = (actualRuleData[i][2] as string).split("\n");
    const line = actualRuleData[i][3] as string;

    sourceIPs.forEach(ip => {
      if (!uniqueServers.has(ip)) {
        uniqueServers.set(ip, new Set());
      }
      uniqueServers.get(ip)?.add(line);
    });

    destinationIPs.forEach(ip => {
      if (!uniqueServers.has(ip)) {
        uniqueServers.set(ip, new Set());
      }
      uniqueServers.get(ip)?.add(line);
    });

    services.forEach(service => uniqueServices.add(service));
  }

  // Zapisz unikalne adresy w arkuszu Servers
  const serversData: (string | number)[][] = [];
  uniqueServers.forEach((lines, ip) => {
    const details = formatIPDetails(ip);
    lines.forEach(line => {
      serversData.push([serversData.length + 1, ip, details, line]);
    });
  });

  sheetServers.getRange("A1:D1").setValues([["ID", "name", "details", "line"]]);
  sheetServers.getRange(`A2:D${serversData.length + 1}`).setValues(serversData);

  // Zapis unikalnych portów w arkuszu Services, zaczynając od określonego wiersza (np. 237)
  const startRowForServices = 237; // Określony wiersz, od którego zaczynamy dodawanie
  const initialServiceID = 16392; // Zaczynamy od 16391 + 1

  const servicesData = Array.from(uniqueServices).map((service, index) => {
    return [initialServiceID + index, service]; // ID oparte na initialServiceID
  });

  // Ustaw nagłówki, jeśli arkusz jest pusty
  if (sheetServices.getUsedRange().getRowCount() === 0) {
    sheetServices.getRange("A1:B1").setValues([["ID", "name"]]);
  }

  // Zapis do arkusza Services
  sheetServices.getRange(`A${startRowForServices}:B${startRowForServices + servicesData.length - 1}`).setValues(servicesData);

  // Zasil arkusz Connections
  const connectionsData = actualRuleData.slice(1).map((row, index) => {
    const sourceIDs = Array.from(row[0].split("\n").map(ip => getIDForIP(ip, serversData)));
    const destinationIDs = Array.from(row[1].split("\n").map(ip => getIDForIP(ip, serversData)));
    const serviceIDs = Array.from(row[2].split("\n").map(service => getIDForService(service, servicesData)));
    return [
      index + 1,
      sourceIDs.join(", "),
      destinationIDs.join(", "),
      serviceIDs.join(", "),
      row[3] // Line
    ];
  });

  sheetConnections.getRange("A1:E1").setValues([["ID", "source IDs", "destination IDs", "service IDs", "line"]]);
  sheetConnections.getRange(`A2:E${connectionsData.length + 1}`).setValues(connectionsData);
}

// Funkcja formatowania szczegółów IP
function formatIPDetails(ip: string): string {
  if (ip.startsWith("net_")) {
    return ip.slice(4).replace(/_/g, "/");
  } else if (isIPAddress(ip)) {
    return ip;
  }
  return ip; // Zwraca IP w przypadku gdy nie jest to poprawny format
}

// Funkcja sprawdzająca, czy ciąg znaków to adres IP
function isIPAddress(ip: string): boolean {
  const ipPattern = /^(\d{1,3}\.){3}\d{1,3}$/;
  return ipPattern.test(ip);
}

// Funkcja do uzyskiwania ID dla adresu IP
function getIDForIP(ip: string, serversData: (string | number)[][]): number {
  const serverIndex = serversData.findIndex(row => row[1] === ip);
  return serverIndex >= 0 ? serverIndex + 1 : 0; // Zwraca ID lub 0, jeśli nie znaleziono
}

// Funkcja do uzyskiwania ID dla usługi
function getIDForService(service: string, servicesData: (string | number)[][]): number {
  const serviceIndex = servicesData.findIndex(row => row[1] === service);
  return serviceIndex >= 0 ? servicesData[serviceIndex][0] as number : 0; // Zwraca ID lub 0, jeśli nie znaleziono
}
