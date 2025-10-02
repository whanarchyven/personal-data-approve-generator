"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { format } from "date-fns";
import { ru } from "date-fns/locale";

type RawRow = {
  role?: string | number;
  full_name?: string | number;
  birth_date?: string | number | Date;
  tutor_name?: string | number;
  tutor_phone?: string | number;
};

type Participant = {
  role: string;
  fullName: string;
  birthDate?: Date | null;
  tutorName?: string;
  tutorPhone?: string;
};

function parseExcelDate(value: unknown): Date | null {
  if (value == null) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value as number);
    if (!d) return null;
    return new Date(d.y, (d.m || 1) - 1, d.d || 1);
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    // Try ISO
    const iso = new Date(trimmed);
    if (!isNaN(iso.getTime())) return iso;
    // Try dd.MM.yyyy
    const m = trimmed.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
    if (m) {
      const day = parseInt(m[1], 10);
      const month = parseInt(m[2], 10) - 1;
      const year = parseInt(m[3].length === 2 ? `20${m[3]}` : m[3], 10);
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
  }
  return null;
}

function isEscort(role: string): boolean {
  const r = String(role || "").trim().toLowerCase();
  return r === "сопровождающий";
}

function formatBirthDate(d?: Date | null): string {
  if (!d) return "";
  return format(d, "dd.MM.yyyy");
}

function formatConsentDate(d?: Date | null): string {
  if (!d) return "";
  const day = format(d, "dd");
  const month = format(d, "LLLL", { locale: ru });
  const year = format(d, "yyyy");
  return `«${day}» ${month} ${year} г.`;
}

function buildChildDoc(p: Participant, orgName: string, dateStr: string) {
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 720,    // 0.5 дюйма = 720 твипов
            right: 720,
            bottom: 720,
            left: 720,
          },
        },
      },
      children: [
        new Paragraph({ text: "СОГЛАСИЕ" }),
        new Paragraph({ text: "РОДИТЕЛЯ (ЗАКОННОГО ПРЕДСТАВИТЕЛЯ) НА ОБРАБОТКУ ПЕРСОНАЛЬНЫХ ДАННЫХ НЕСОВЕРШЕННОЛЕТНЕГО" }),
        new Paragraph({ text: "" }),
        new Paragraph({ text: `Я, ${p.tutorName || ""},` }),
        new Paragraph({ text: `являюсь законным представителем несовершеннолетнего ${p.fullName}` }),
        new Paragraph({ text: `дата рождения: ${formatBirthDate(p.birthDate)},` }),
        new Paragraph({ text: `номер мобильного телефона законного представителя: ${p.tutorPhone || ""},` }),
        new Paragraph({ text: "" }),
        new Paragraph({ text: `в соответствии с ч. 4 ст. 9 Федерального закона от 27.07.2006 № 152-ФЗ «О персональных данных» даю свое согласие на обработку в ${orgName} персональных данных несовершеннолетнего, относящихся исключительно к перечисленным ниже категориям персональных данных: фамилия, имя, отчество; год, месяц, дата рождения; контактный номер телефона для связи; наименование образовательной организации, класс, фото и видеосъемку, а также их публикацию в социальных сетях, на сайте и СМИ. Я даю согласие на использование персональных данных несовершеннолетнего в целях организации и проведения культурно-просветительских программ для школьников, а также формирования отчетности по реализации национального проекта «Семья» для последующей обработки ${orgName} персональных данных следующими способами: накопление, хранение, передача в Министерство культуры Российской Федерации, удаление и уничтожение на бумажных и/или электронных носителях. Настоящее согласие предоставляется мной на осуществление действий в отношении персональных данных несовершеннолетнего, которые необходимы для достижения указанных выше целей, включая (без ограничения) сбор, систематизацию, накопление, хранение, уточнение (обновление, изменение), использование, передачу третьим лицам для осуществления действий по обмену информацией (Общероссийской общественно-государственной организации «Российский фонд культуры, Министерству Культуры Российской Федерации), обезличивание, блокирование персональных данных, а также осуществление любых иных действий, предусмотренных действующим законодательством Российской Федерации. Я проинформирован, что ${orgName} гарантирует обработку персональных данных несовершеннолетнего в соответствии с действующим законодательством Российской Федерации как неавтоматизированным, так и автоматизированным способами. Данное согласие действует до достижения целей обработки персональных данных или в течение срока хранения информации. Данное согласие может быть отозвано в любой момент по моему письменному заявлению. Я подтверждаю, что, давая такое согласие, я действую по собственной воле и в интересах несовершеннолетнего.` }),
        new Paragraph({ text: "" }),
        new Paragraph({ text: `${dateStr}                         __________________ /${p.tutorName || ""}/` }),
      ]
    }]
  });
  return doc;
}

function buildEscortDoc(p: Participant, orgName: string, dateStr: string) {
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 720,    // 0.5 дюйма = 720 твипов
            right: 720,
            bottom: 720,
            left: 720,
          },
        },
      },
      children: [
        new Paragraph({ text: "СОГЛАСИЕ", heading: HeadingLevel.HEADING_2 }),
        new Paragraph({ text: "НА ОБРАБОТКУ ПЕРСОНАЛЬНЫХ ДАННЫХ" }),
        new Paragraph({ text: "" }),
        new Paragraph({ text: `Я, ${p.fullName},` }),
        new Paragraph({ text: `дата рождения: ${formatBirthDate(p.birthDate)},` }),
        new Paragraph({ text: `номер мобильного телефона: ${p.tutorPhone || ""},` }),
        new Paragraph({ text: "" }),
        new Paragraph({ text: `в соответствии с ч. 4 ст. 9 Федерального закона от 27.07.2006 № 152-ФЗ «О персональных данных» даю свое согласие на обработку в ${orgName} моих персональных данных, относящихся исключительно к перечисленным ниже категориям персональных данных: фамилия, имя, отчество; год, месяц, дата рождения; контактный номер телефона для связи; наименование образовательной организации, фото и видеосъемку, а также их публикацию в социальных сетях, на сайте и СМИ. Я даю согласие на использование моих персональных данных в целях организации и проведения культурно-просветительских программ для школьников, а также формирования отчетности по реализации национального проекта «Семья» для последующей обработки ${orgName} моих персональных данных следующими способами: накопление, хранение, передача в Министерство культуры Российской Федерации, удаление и уничтожение на бумажных и/или электронных носителях. Настоящее согласие предоставляется мной на осуществление действий в отношении моих персональных данных, которые необходимы для достижения указанных выше целей, включая (без ограничения) сбор, систематизацию, накопление, хранение, уточнение (обновление, изменение), использование, передачу третьим лицам для осуществления действий по обмену информацией (Общероссийской общественно-государственной организации «Российский фонд культуры, Министерству Культуры Российской Федерации), обезличивание, блокирование моих персональных данных, а также осуществление любых иных действий, предусмотренных действующим законодательством Российской Федерации. Я проинформирован, что ${orgName} гарантирует обработку моих персональных данных в соответствии с действующим законодательством Российской Федерации как неавтоматизированным, так и автоматизированным способами. Данное согласие действует до достижения целей обработки моих персональных данных или в течение срока хранения информации. Данное согласие может быть отозвано в любой момент по моему письменному заявлению. Я подтверждаю, что, давая такое согласие, я действую по собственной воле.` }),
        new Paragraph({ text: "" }),
        new Paragraph({ text: `${dateStr}                         __________________ /${p.fullName}/` }),
      ]
    }]
  });
  return doc;
}


export default function Home() {
  const [orgName, setOrgName] = useState("");
  const [dateStr, setDateStr] = useState(""); // from <input type="date">
  const [rawRows, setRawRows] = useState<Participant[]>([]);

  const consentDate = useMemo(() => {
    if (!dateStr) return null;
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? null : d;
  }, [dateStr]);

  const escorts = useMemo(() => rawRows.filter((r) => isEscort(r.role)), [rawRows]);
  const children = useMemo(() => rawRows.filter((r) => !isEscort(r.role)), [rawRows]);

  async function onFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    console.log("[DEBUG 1] Файл выбран:", file?.name, file?.type);
    
    if (!file) return;
    let rows: RawRow[] = [];
    
    if (file.name.toLowerCase().endsWith(".json") || file.type === "application/json") {
      console.log("[DEBUG 2] Читаем как JSON");
      const text = await file.text();
      console.log("[DEBUG 3] Текст JSON (первые 500 символов):", text.substring(0, 500));
      try {
        const data = JSON.parse(text);
        console.log("[DEBUG 4] JSON распарсен, тип:", typeof data, "Array?", Array.isArray(data));
        if (Array.isArray(data)) {
          rows = data as RawRow[];
          console.log("[DEBUG 5] JSON массив, длина:", rows.length);
        }
      } catch (err) {
        console.error("[DEBUG ERROR] JSON parse error", err);
        return;
      }
    } else {
      console.log("[DEBUG 2] Читаем как Excel");
      const buf = await file.arrayBuffer();
      console.log("[DEBUG 3] ArrayBuffer размер:", buf.byteLength);
      
      const wb = XLSX.read(buf, { type: "array" });
      console.log("[DEBUG 4] Workbook прочитан, листы:", wb.SheetNames);
      
      const sheetName = wb.SheetNames[0];
      console.log("[DEBUG 5] Выбран лист:", sheetName);
      
      const ws = wb.Sheets[sheetName];
      console.log("[DEBUG 6] Worksheet объект:", ws);
      console.log("[DEBUG 7] Worksheet !ref (диапазон):", ws['!ref']);
      
      rows = XLSX.utils.sheet_to_json<RawRow>(ws, { defval: "" });
      console.log("[DEBUG 8] sheet_to_json завершен, строк:", rows.length);
    }

    console.log("[DEBUG 9] Всего строк итого:", rows.length);
    
    // Преобразуем в Participant
    const participants: Participant[] = rows
      .map((row) => ({
        role: String(row.role || ""),
        fullName: String(row.full_name || ""),
        birthDate: parseExcelDate(row.birth_date),
        tutorName: String(row.tutor_name || ""),
        tutorPhone: String(row.tutor_phone || ""),
      }))
      .filter((p) => {
        if (!p.fullName) return false;
        const r = p.role.toLowerCase();
        if (r === "итого") return false;
        if (/^\d+$/.test(p.fullName)) return false;
        return true;
      });
    
    console.log("[DEBUG 10] Обработано участников:", participants.length);
    console.log("[DEBUG 11] Сопровождающих:", participants.filter(p => isEscort(p.role)).length);
    console.log("[DEBUG 12] Детей:", participants.filter(p => !isEscort(p.role)).length);

    setRawRows(participants);
  }

  function sanitizeFileName(name: string): string {
    return name.replace(/[\\/:*?"<>|]/g, "_");
  }

  async function onGenerate() {
    if (!consentDate || !orgName || rawRows.length === 0) return;
    
    const zip = new JSZip();
    const formattedConsentDate = formatConsentDate(consentDate);

    const allDocs = await Promise.all(
      rawRows.map(async (p) => {
        const doc = isEscort(p.role)
          ? buildEscortDoc(p, orgName, formattedConsentDate)
          : buildChildDoc(p, orgName, formattedConsentDate);
        const blob = await Packer.toBlob(doc);
        return { name: `${sanitizeFileName(p.fullName)}.docx`, blob };
      })
    );

    for (const { name, blob } of allDocs) {
      zip.file(name, blob);
    }

    const zipBlob = await zip.generateAsync({ type: "blob" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(zipBlob);
    a.download = `согласия_${format(new Date(), "yyyyMMdd_HHmmss")}.zip`;
    document.body.appendChild(a);
    a.click();
    URL.revokeObjectURL(a.href);
    a.remove();
  }

  const canGenerate = orgName.trim().length > 0 && !!consentDate && rawRows.length > 0;

  return (
    <div className="min-h-screen w-full px-6 py-10 flex flex-col gap-8">
      <div className="max-w-3xl w-full mx-auto flex flex-col gap-6">
        <h1 className="text-2xl font-semibold">Генератор согласий на обработку ПД</h1>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="flex flex-col gap-2 md:col-span-2">
            <label className="text-sm text-gray-600">Наименование организации</label>
            <input
              type="text"
              value={orgName}
              onChange={(e) => setOrgName(e.target.value)}
              placeholder="Например: МБОУ Школа №1"
              className="border rounded px-3 py-2"
            />
          </div>
          <div className="flex flex-col gap-2">
            <label className="text-sm text-gray-600">Дата согласий</label>
            <input
              type="date"
              value={dateStr}
              onChange={(e) => setDateStr(e.target.value)}
              className="border rounded px-3 py-2"
            />
          </div>
        </div>

        <div className="flex flex-col gap-2">
          <label className="text-sm text-gray-600">Загрузить файл с участниками (Excel/JSON)</label>
          <input type="file" accept=".xlsx,.xls,.json" onChange={onFileChange} />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div>
            <h2 className="font-medium mb-2">Сопровождающие ({escorts.length})</h2>
            <ul className="space-y-1">
              {escorts.map((p, i) => (
                <li key={`e-${i}`} className="text-sm">
                  {p.fullName}
                </li>
              ))}
            </ul>
          </div>
          <div>
            <h2 className="font-medium mb-2">Дети ({children.length})</h2>
            <ul className="space-y-1">
              {children.map((p, i) => (
                <li key={`c-${i}`} className="text-sm">
                  {p.fullName} {p.tutorName ? `(представитель: ${p.tutorName})` : ""}
                </li>
              ))}
            </ul>
          </div>
        </div>

        <div className="pt-2">
          <button
            onClick={onGenerate}
            disabled={!canGenerate}
            className={`px-4 py-2 rounded font-medium ${
              canGenerate
                ? "bg-black text-white hover:bg-gray-800"
                : "bg-gray-200 text-gray-500 cursor-not-allowed"
            }`}
          >
            Сгенерировать согласия (ZIP)
          </button>
        </div>
      </div>
    </div>
  );
}
