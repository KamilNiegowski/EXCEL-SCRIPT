function main(workbook: ExcelScript.Workbook) {
  const sheetActualRule = workbook.getWorksheet("ACTUAL RULE");
  const sheetServers = workbook.getWorksheet("Servers") || workbook.addWorksheet("Servers");
  const sheetServices = workbook.getWorksheet("Services") || workbook.addWorksheet("Services");
  const sheetConnections = workbook.getWorksheet("Connections") || workbook.addWorksheet("Connections");

  // Pobranie danych z ACTUAL RULE
  const actualRuleData = sheetActualRule.getUsedRange().getValues();

  // Zbiór unikalnych adresów IP i portów
  const uniqueServers = new Map<string, Set<string>>(); // Klucz: IP/nazwa, wartość: Zbiór linii
  const uniqueServices = new Map<string, number>(); // Klucz: usługa (port), wartość: unikalne ID

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

      services.forEach(service => {
          const lowerService = service.toLowerCase();
          if (!uniqueServices.has(lowerService)) {
              uniqueServices.set(lowerService, 0); // Zainicjuj usługę, jeśli jeszcze jej nie ma
          }
      });
  }

  // Uzupełnianie arkusza Servers
  const serversData: (string | number)[][] = [];
  uniqueServers.forEach((lines, ip) => {
      const details = formatIPDetails(ip);
      lines.forEach(line => {
          serversData.push([serversData.length + 1, ip, details, line]);
      });
  });

      sheetServers.getRange("A1:D1").setValues([["ID", "name", "details", "line"]]);
      sheetServers.getRange(`A2:D${serversData.length + 1}`).setValues(serversData);


  // Sprawdzenie ostatniego ID w arkuszu Services
  const lastServiceID = getLastID(sheetServices);
  const existingServices = getExistingEntries(sheetServices);
  const servicesData: (string | number)[][] = [];
  let nextServiceID = lastServiceID + 1; // Rozpocznij od najwyższego ID + 1

  // Dodaj nowe usługi do arkusza Services
  uniqueServices.forEach((_, service) => {
      if (!existingServices.has(service.toLowerCase())) {
          servicesData.push([nextServiceID++, service]);
      }
  });

  if (servicesData.length > 0) {
      sheetServices.getRange("A1:B1").setValues([["ID", "name"]]);
      sheetServices.getRange(`A${lastServiceID + 2}:B${lastServiceID + servicesData.length + 1}`).setValues(servicesData);
  }

  // Uzupełnianie arkusza Connections
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
  return ip;
}

// Funkcja sprawdzająca, czy ciąg znaków to adres IP
function isIPAddress(ip: string): boolean {
  const ipPattern = /^(\d{1,3}\.){3}\d{1,3}$/;
  return ipPattern.test(ip);
}

// Funkcja do uzyskiwania ID dla adresu IP
function getIDForIP(ip: string, serversData: (string | number)[][]): number {
  const serverIndex = serversData.findIndex(row => row[1].toLowerCase() === ip.toLowerCase());
  return serverIndex >= 0 ? serversData[serverIndex][0] as number : 0;
}

// Funkcja do uzyskiwania ID dla usługi
function getIDForService(service: string, servicesData: (string | number)[][]): number {
  const serviceIndex = servicesData.findIndex(row => row[1].toLowerCase() === service.toLowerCase());
  return serviceIndex >= 0 ? servicesData[serviceIndex][0] as number : 0;
}

// Funkcja do uzyskiwania następnego ID w arkuszu
function getNextID(sheet: ExcelScript.Worksheet): number {
  const usedRange = sheet.getUsedRange();
  const lastRow = usedRange ? usedRange.getRowCount() : 0;

  // Jeśli arkusz jest pusty, zacznij od 1
  if (lastRow <= 1) return 1; // Zakłada, że pierwszy wiersz to nagłówki

  // Sprawdź ostatnie ID w kolumnie A
  let lastID = 0;
  for (let i = 2; i <= lastRow; i++) { // Zaczynamy od 2, aby pominąć nagłówki
      const idValue = sheet.getRange(`A${i}`).getValue();
      if (typeof idValue === "number") {
          lastID = Math.max(lastID, idValue); // Zapisz największe ID
      }
  }

  return lastID + 1; // Zwróć następne ID
}

// Funkcja do pobierania ostatniego ID w arkuszu Services
function getLastID(sheet: ExcelScript.Worksheet): number {
  const usedRange = sheet.getUsedRange();
  const lastRow = usedRange.getLastRow();
  const lastID = sheet.getRange(`A${lastRow.getRowIndex() + 1}`).getValue() as number;
  return isNaN(lastID) ? 0 : lastID;
}

// Funkcja do pobierania istniejących usług z arkusza
function getExistingEntries(sheet: ExcelScript.Worksheet): Set<string> {
  const entries = new Set<string>();
  const data = sheet.getUsedRange().getValues();
  for (let i = 1; i < data.length; i++) {
      entries.add((data[i][1] as string).toLowerCase());
  }
  return entries;
}
