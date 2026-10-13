import dayjs from 'dayjs';
import * as XLSX from 'xlsx';

const locatorSheetName = 'РЛС';
const lightSheetName = 'СТП ЗРК МД';
const heavySheetName = 'СТП ППО';

const convertExcelDateToJsDate = (xTime: number) => {
  const { y, m, d, H, M } = XLSX.SSF.parse_date_code(xTime);

  return dayjs(new Date(y, m - 1, d, H, M));
};

const dmsToDecimal = (coordinate: number) => {
  const value = String(coordinate).padStart(6, '0');

  const degrees = Number(value.slice(0, value.length - 4));
  const minutes = Number(value.slice(-4, -2));
  const seconds = Number(value.slice(-2));

  return degrees + minutes / 60 + seconds / 3600;
};

const file = (features: any[]) => {
  const base = {
    type: 'FeatureCollection',
    features,
  };
  const blob = new Blob([JSON.stringify(base, null, 2)], {
    type: 'application/geo+json;charset=utf-8',
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');

  link.href = url;
  link.download = 'filename';
  document.body.appendChild(link);
  link.click();
  link.remove();

  URL.revokeObjectURL(url);
};

const CONFIG = {
  locator: {
    name: 'A',
    location: 'B',
    region: 'C',
    lat: 'D',
    lon: 'E',
  },
  heavy: {
    name: 'A',
    date: 'B',
    location: 'C',
    region: 'D',
    lat: 'E',
    lon: 'F',
  },
};

export class ExcelReader {
  private static options: XLSX.ParsingOptions = {
    type: 'array',
    cellStyles: true,
    sheetStubs: true,
    FS: '||||',
  };

  static locators = [];

  static read(data: Uint8Array) {
    const { Sheets } = XLSX.read(data, this.options);

    const locatorRef = Sheets[locatorSheetName]['!ref']!;
    const lightRef = Sheets[lightSheetName]['!ref']!;
    const heavyRef = Sheets[heavySheetName]['!ref']!;
    const locatorData = Sheets[locatorSheetName];
    const heavyData = Sheets[heavySheetName];
    const lightData = Sheets[lightSheetName];

    const {
      e: { r: locatorRows },
    } = XLSX.utils.decode_range(locatorRef);

    const {
      e: { r: lightRows },
    } = XLSX.utils.decode_range(lightRef);

    const {
      e: { r: heavyRows },
    } = XLSX.utils.decode_range(heavyRef);

    for (let i = 1; i <= locatorRows; i++) {
      const sidc = '10061500002203000000';
      const nameKey = `${CONFIG.locator.name}${i}`;
      const latKey = `${CONFIG.locator.lat}${i}`;
      const lonKey = `${CONFIG.locator.lon}${i}`;
      const locationKey = `${CONFIG.locator.location}${i}`;
      const regionKey = `${CONFIG.locator.region}${i}`;

      if (!locatorData[nameKey]?.v) {
        continue;
      }

      const locator = {
        type: 'Feature',
        properties: {
          sidc,
          name: locatorData[nameKey]?.v,
          observation_datetime: Date.now(),
          staff_comments: `${locatorData[locationKey]?.v}, ${locatorData[regionKey]?.v}`,
          comments: [
            `${locatorData[regionKey]?.v}`,
            locatorData[locationKey]?.v,
          ],
        },
        geometry: {
          type: 'Point',
          coordinates: [
            this.dmsToDecimal(locatorData[lonKey]?.v),
            this.dmsToDecimal(locatorData[latKey]?.v),
          ],
        },
      };

      this.locators.push(locator);
    }
    for (let i = 1; i <= lightRows; i++) {
      const sidc = '10061500001111000000';
      const nameKey = `${CONFIG.locator.name}${i}`;
      const latKey = `${CONFIG.locator.lat}${i}`;
      const lonKey = `${CONFIG.locator.lon}${i}`;
      const locationKey = `${CONFIG.locator.location}${i}`;
      const regionKey = `${CONFIG.locator.region}${i}`;

      if (!lightData[nameKey]?.v) {
        continue;
      }

      const locator = {
        type: 'Feature',
        properties: {
          sidc,
          name: lightData[nameKey]?.v,
          observation_datetime: Date.now(),
          staff_comments: `${lightData[locationKey]?.v}, ${lightData[regionKey]?.v}`,
          comments: [`${lightData[regionKey]?.v}`, lightData[locationKey]?.v],
        },
        geometry: {
          type: 'Point',
          coordinates: [
            this.dmsToDecimal(lightData[lonKey]?.v),
            this.dmsToDecimal(lightData[latKey]?.v),
          ],
        },
      };

      this.locators.push(locator);
    }
    for (let i = 1; i <= heavyRows; i++) {
      const sidc = '10061500001111000000';
      const nameKey = `${CONFIG.heavy.name}${i}`;
      const latKey = `${CONFIG.heavy.lat}${i}`;
      const lonKey = `${CONFIG.heavy.lon}${i}`;
      const dateKey = `${CONFIG.heavy.date}${i}`;
      const locationKey = `${CONFIG.heavy.location}${i}`;
      const regionKey = `${CONFIG.heavy.region}${i}`;

      const dateValue = heavyData[dateKey]?.v;

      if (!heavyData[nameKey]?.v) {
        continue;
      }

      const locator = {
        type: 'Feature',
        properties: {
          sidc,
          name: heavyData[nameKey]?.v,
          observation_datetime: convertExcelDateToJsDate(dateValue).format(
            'YYYY-MM-DDTHH:mm:ss'
          ),
          staff_comments: `${heavyData[locationKey]?.v}, ${heavyData[regionKey]?.v}`,
          comments: [`${heavyData[regionKey]?.v}`, heavyData[locationKey]?.v],
        },
        geometry: {
          type: 'Point',
          coordinates: [
            this.dmsToDecimal(heavyData[lonKey]?.v),
            this.dmsToDecimal(heavyData[latKey]?.v),
          ],
        },
      };

      this.locators.push(locator);
    }
  }

  static isDownloadawailable = () => {
    return !!this.locators.length;
  };

  static file = () => {
    const base = {
      type: 'FeatureCollection',
      features: this.locators,
    };
    const blob = new Blob([JSON.stringify(base, null, 2)], {
      type: 'application/geo+json;charset=utf-8',
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = 'filename';
    document.body.appendChild(link);
    link.click();
    link.remove();

    URL.revokeObjectURL(url);
  };

  private static dmsToDecimal = (coordinate: number) => {
    const value = String(coordinate).padStart(6, '0');

    const degrees = Number(value.slice(0, value.length - 4));
    const minutes = Number(value.slice(-4, -2));
    const seconds = Number(value.slice(-2));

    return degrees + minutes / 60 + seconds / 3600;
  };
}
