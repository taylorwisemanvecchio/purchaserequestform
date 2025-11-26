import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  DayOfWeek,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Stack,
  ChoiceGroup,
  IChoiceGroupOption,
  ComboBox,
  IComboBoxOption,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { IPurchaseRequestFormProps } from "./IPurchaseRequestFormProps";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";
import "@pnp/sp/profiles"; // <-- for sp.profiles.myProperties()

const LIST_NAMES = {
  requests: "PO_Requests",
  lines: "PO_LineItems",
  vendors: "Vendors",
} as const;

const REQUESTS_VENDOR_LOOKUP_INTERNAL_NAME = "Vendor";
const REQUESTS_LINEITEMS_JSON_FIELD = "LineItemsJson";

// ---------------- Department Manager Mapping ----------------
type DeptRow = { office: string; department: string; email: string; name: string };

const DEPT_MANAGER_MAP: DeptRow[] = [
  { office: "MTL",       department: "Corporate",      email: "kane@taylorwiseman.com",      name: "Patrick Kane" },
  { office: "MLT",       department: "Engineering",    email: "vecchio@taylorwiseman.com",   name: "Gary Vecchio" },
  { office: "MLT",       department: "Survey",         email: "previtera@taylorwiseman.com", name: "Sam Previtera" },
  { office: "MTL",       department: "Admin",          email: "kane@taylorwiseman.com",      name: "Patrick Kane" },
  { office: "MTL",       department: "Transportation", email: "dmoore@taylorwiseman.com",    name: "Dennis Moore" },
  { office: "MLT",       department: "SUE",            email: "taylor@taylorwiseman.com",    name: "Adam Taylor" },
  { office: "MTL",       department: "Geospatial",     email: "kane@taylorwiseman.com",      name: "Patrick Kane" },
  { office: "Blue Bell", department: "Engineering",    email: "thompson@taylorwiseman.com",  name: "Mark Thompson" },
  { office: "Blue Bell", department: "Survey",         email: "nowicki@taylorwiseman.com",   name: "Dave Nowicki" },
  { office: "Blue Bell", department: "Environmental",  email: "maher@taylorwiseman.com",     name: "Tom Maher" },
  { office: "Apex",      department: "Engineering",    email: "pearsall@taylorwiseman.com",  name: "Jim Pearsall" },
  { office: "Apex",      department: "Survey",         email: "pearsall@taylorwiseman.com",  name: "Jim Pearsall" },
  { office: "Apex",      department: "Corporate",      email: "pearsall@taylorwiseman.com",  name: "Jim Pearsall" },
  { office: "Apex",      department: "SUE",            email: "pacellat@taylorwiseman.com",  name: "Tom Pacella" },
  { office: "Charlotte", department: "Engineering",    email: "pearsall@taylorwiseman.com",  name: "Jim Pearsall" },
  { office: "Charlotte", department: "Survey",         email: "pearsall@taylorwiseman.com",  name: "Jim Pearsall" },
  { office: "Charlotte", department: "Corporate",      email: "pearsall@taylorwiseman.com",  name: "Jim Pearsall" },
  { office: "Charlotte", department: "SUE",            email: "pacellat@taylorwiseman.com",  name: "Tom Pacella" },
];

const norm = (s: unknown): string =>
  String(s ?? "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim().replace(/\s+/g, " ");

const pickDeptManager = (officeRaw: string, deptRaw: string): DeptRow | undefined => {
  const office = norm(officeRaw);
  const dept = norm(deptRaw);

  const exact = DEPT_MANAGER_MAP.find(r => norm(r.office) === office && norm(r.department) === dept);
  if (exact) return exact;

  const byDept = DEPT_MANAGER_MAP.find(r => norm(r.department) === dept);
  if (byDept) return byDept;

  const byOffice = DEPT_MANAGER_MAP.find(r => norm(r.office) === office);
  if (byOffice) return byOffice;

  return undefined;
};

const DEPT_MANAGER_OPTIONS: IDropdownOption[] = Array.from(
  new Map(DEPT_MANAGER_MAP.map(r => [r.email, r])).values()
).map(r => ({ key: r.email, text: `${r.name} — ${r.email}` }));

// ---------------- /Department Manager Mapping ----------------

const EXPENSE_CATEGORIES: IDropdownOption[] = [
  { key: "Training", text: "Training" },
  { key: "Meetings-Conventions", text: "Meetings/Conventions" },
  { key: "CareerFair", text: "Career Fair" },
  { key: "CharitableEvents", text: "Charitable Events" },
  { key: "EmployeeRecreation", text: "Employee Recreation" },
  { key: "Advertising", text: "Advertising" },
  { key: "PromotionalItems-Apparel", text: "Promotional Items / Apparel" },
  { key: "ComputerSupplies-Equipment-Software", text: "Computer Supplies / Equipment / Software" },
  { key: "OfficeSupplies-Equipment", text: "Office Supplies / Equipment" },
  { key: "FieldSupplies-Equipment", text: "Field Supplies / Equipment" },
  { key: "Maps-Deeds-Certified Owners", text: "Maps / Deeds / Certified Owners" },
];

const COST_CENTERS: IDropdownOption[] = [
  { key: "1030-Corporate-Expense-Overhead", text: "1030 Corporate Expense Overhead" },
  { key: "2010-MtLaurel-Engineering", text: "2010 Mt. Laurel – Engineering" },
  { key: "2020-MtLaurel-Survey", text: "2020 Mt. Laurel – Survey" },
  { key: "2020-Vineland", text: "2020 Vineland" },
  { key: "2030-MtLaurel-Admin", text: "2030 Mt. Laurel – Admin" },
  { key: "2040-MtLaurel-Transportation", text: "2040 Mt. Laurel – Transportation" },
  { key: "2050-MtLaurel-SUE", text: "2050 Mt. Laurel – SUE" },
  { key: "2090-MtLaurel-Geospatial", text: "2090 Mt. Laurel – Geospatial" },
  { key: "5010-BlueBell-Engineering", text: "5010 Blue Bell – Engineering" },
  { key: "5020-BlueBell-Survey", text: "5020 Blue Bell – Survey" },
  { key: "5030-BlueBell-Admin", text: "5030 Blue Bell – Admin" },
  { key: "5030-Bethlehem", text: "5030 Bethlehem" },
  { key: "5030-Bethlehem-Admin", text: "5030 Bethlehem – Admin" },
  { key: "5070-BlueBell-Enviro", text: "5070 Blue Bell – Enviro" },
  { key: "6020-Apex-Survey", text: "6020 Apex - Survey" },
  { key: "6030-Apex-Admin", text: "6030 Apex – Admin" },
  { key: "6050-Apex-SUE", text: "6050 Apex – SUE" },
  { key: "6130-Charlotte-Admin", text: "6130 Charlotte – Admin" },
  { key: "6150-Charlotte-SUE", text: "6150 Charlotte – SUE" },
];

const PAYMENT_METHODS: IChoiceGroupOption[] = [
  { key: "Credit Card", text: "Credit Card" },
  { key: "Check Request", text: "Check Request" },
  { key: "Invoice", text: "Invoice" },
];

interface ILineItem {
  description: string;
  qty: string;        // store as string to preserve user input states
  unitCost: string;   // store as string to preserve user input states
  costCenter?: string;
}
interface IVendorItem {
  Id: number;
  Title: string;
  Address?: string;
  Website?: string;
}
interface IVendorOption extends IComboBoxOption {
  data?: { id?: number; address?: string; website?: string };
}
interface IPrevLineItem {
  Id: number;
  ItemName?: string;
  UnitPrice?: number;
  CostCenter?: string;
}
type ReturnMethod = "Mail" | "Return to Requestor";

const MAX_LINES = Number.POSITIVE_INFINITY;
const currency = (n: number): string =>
  isFinite(n) ? n.toLocaleString(undefined, { style: "currency", currency: "USD" }) : "$";

const hasMessage = (e: unknown): e is { message: string } =>
  typeof e === "object" && e !== null && "message" in e && typeof (e as { message: unknown }).message === "string";
const messageFromUnknownError = (e: unknown): string => {
  if (hasMessage(e)) return e.message;
  try { return JSON.stringify(e); } catch { return String(e); }
};

// ---- Helpers (no `any`, no external types needed) ----
type PnpItemAddResult = {
  data?: { Id?: number; ID?: number; id?: number };
  Id?: number;
  ID?: number;
  id?: number;
};

const getItemIdFromAddResult = (ar: PnpItemAddResult): number | undefined => {
  // Check all possible locations where PnP JS might return the ID
  // depending on version and configuration
  const d = ar?.data;
  return d?.Id ?? d?.ID ?? d?.id ?? ar?.Id ?? ar?.ID ?? ar?.id;
};

// Escape single quotes for OData filter strings
const spSingleQuote = (s: string): string => String(s || "").replace(/'/g, "''");

// Parse a numeric string safely; allow "" / "." etc. by returning 0 until valid
const toNum = (s: string): number => {
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
};

// Simple validator for up to 2 decimal places
const isMoneyLike = (s: string): boolean => /^\d*(?:\.\d{0,2})?$/.test(s);

// ------------------------------------------------------

const PurchaseRequestForm: React.FC<IPurchaseRequestFormProps> = (props) => {
  const { sp } = props;

  // ---------- Refs ----------
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // ---------- Form State ----------
  const [expenseCategory, setExpenseCategory] = useState<string | undefined>();
  const [vendorOptions, setVendorOptions] = useState<IVendorOption[]>([]);
  const [vendorKey, setVendorKey] = useState<string | number | undefined>();
  const [addingVendor, setAddingVendor] = useState<boolean>(false);
  const [newVendorName, setNewVendorName] = useState<string>("");
  const [newVendorAddr, setNewVendorAddr] = useState<string>("");
  const [newVendorWebsite, setNewVendorWebsite] = useState<string>("");

  const [requestDate, setRequestDate] = useState<Date | undefined>(undefined);

  const [lineItems, setLineItems] = useState<ILineItem[]>([{ description: "", qty: "", unitCost: "" }]);

  const [files, setFiles] = useState<FileList | null>(null);
  const [salesTax, setSalesTax] = useState<string>("");

  const [paymentMethod, setPaymentMethod] = useState<string | undefined>();

  // Department Manager (selected)
  const [deptMgrEmail, setDeptMgrEmail] = useState<string | undefined>(undefined);
  const [deptMgrName, setDeptMgrName] = useState<string | undefined>(undefined);

  // Status
  const [busy, setBusy] = useState<boolean>(false);
  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);

  // Previous items
  const [prevItemOptions, setPrevItemOptions] = useState<IComboBoxOption[]>([]);
  const [selectedPrevKeys, setSelectedPrevKeys] = useState<readonly (string | number)[]>([]);

  // Check Request fields
  const [jobNo, setJobNo] = useState<string>("");
  const [taxId, setTaxId] = useState<string>("");
  const [telephone, setTelephone] = useState<string>("");
  const [checkAmount, setCheckAmount] = useState<string>("");
  const [dateNeeded, setDateNeeded] = useState<Date | undefined>(undefined);
  const [reason, setReason] = useState<string>("");
  const [returnMethod, setReturnMethod] = useState<ReturnMethod | undefined>(undefined);

  // Sequence guard to avoid race conditions when rapidly switching vendors
  const loadSeq = useRef(0);

  // ---------- Derived ----------
  const subTotal = useMemo<number>(() => {
    return lineItems.reduce((acc, li) => acc + toNum(li.qty) * toNum(li.unitCost), 0);
  }, [lineItems]);

  const total = useMemo<number>(() => subTotal + toNum(salesTax), [subTotal, salesTax]);

  const isOver200 = total >= 200;

  const selectedVendor = useMemo(() => {
    return vendorOptions.find(o => o.key === vendorKey) as IVendorOption | undefined;
  }, [vendorOptions, vendorKey]);

  // ---------- Effects ----------
  useEffect((): void => {
    const loadVendors = async (): Promise<void> => {
      setBusy(true);
      try {
        const items = (await sp.web.lists
          .getByTitle(LIST_NAMES.vendors)
          .items.select("Id,Title,Address,Website")
          .top(500)()) as IVendorItem[];

        const opts: IVendorOption[] = items.map((v) => ({
          key: v.Id,
          text: v.Address ? `${v.Title} — ${v.Address}` : v.Title,
          data: { id: v.Id, address: v.Address, website: v.Website },
        }));
        setVendorOptions(opts);
      } catch (e) {
        setMessage({ type: MessageBarType.error, text: `Failed to initialize: ${messageFromUnknownError(e)}` });
      } finally {
        setBusy(false);
      }
    };

    loadVendors().catch((e) =>
      setMessage({ type: MessageBarType.error, text: `Init failed: ${messageFromUnknownError(e)}` })
    );
  }, [sp]);

  // Load previous items **by selected vendor** (two-step: headers -> lines) with race guard
  useEffect((): void => {
    setPrevItemOptions([]);
    setSelectedPrevKeys([]);

    if (!vendorKey) return;

    const seq = ++loadSeq.current;

    const loadPrevItemsForVendor = async (): Promise<void> => {
      setBusy(true);
      try {
        const headers = await sp.web.lists
          .getByTitle(LIST_NAMES.requests)
          .items.select("Id")
          .filter(`${REQUESTS_VENDOR_LOOKUP_INTERNAL_NAME}Id eq ${vendorKey}`)
          .top(500)();

        if (loadSeq.current !== seq) return;

        const headerIds: number[] = (headers || [])
          .map((h: { Id: number }) => h.Id)
          .filter((id) => typeof id === "number");

        if (!headerIds.length) {
          if (loadSeq.current === seq) setPrevItemOptions([]);
          return;
        }

        const chunkSize = 20;
        const allLines: IPrevLineItem[] = [];
        for (let i = 0; i < headerIds.length; i += chunkSize) {
          const ids = headerIds.slice(i, i + chunkSize);
          const orFilter = ids.map((id) => `PO_IDId eq ${id}`).join(" or ");
          const lines = await sp.web.lists
            .getByTitle(LIST_NAMES.lines)
            .items.select("Id,ItemName,UnitPrice,CostCenter,PO_IDId")
            .filter(orFilter)
            .top(500)();
          allLines.push(...(lines as IPrevLineItem[]));

          if (loadSeq.current !== seq) return;
        }

        const keyMap = new Map<string, IPrevLineItem>();
        for (const it of allLines) {
          const name = (it.ItemName || "").trim();
          const price = Number(it.UnitPrice || 0);
          const cc = it.CostCenter || "";
          const key = `${name}|${price}|${cc}`;
          if (name && !keyMap.has(key)) keyMap.set(key, it);
        }

        if (loadSeq.current !== seq) return;

        const options: IComboBoxOption[] = Array.from(keyMap.entries()).map(([k, v]) => ({
          key: k,
          text: `${v.ItemName || "(no name)"} — ${currency(Number(v.UnitPrice || 0))}${v.CostCenter ? ` — ${v.CostCenter}` : ""}`,
          data: { itemName: v.ItemName || "", unitPrice: Number(v.UnitPrice || 0), costCenter: v.CostCenter || "" },
        }));
        setPrevItemOptions(options);
      } catch (e) {
        if (loadSeq.current === seq) {
          setMessage({
            type: MessageBarType.error,
            text: `Failed to load previous items: ${messageFromUnknownError(e)}`,
          });
        }
      } finally {
        if (loadSeq.current === seq) setBusy(false);
      }
    };

    loadPrevItemsForVendor().catch((e) => {
      if (loadSeq.current === seq) {
        setMessage({ type: MessageBarType.error, text: `Prev items failed: ${messageFromUnknownError(e)}` });
      }
    });
  }, [vendorKey, sp]);

  // Preselect Department Manager from current user's Office & Department (O365 profile)
  type ProfileEntry = { Key?: string; Value?: string };

  useEffect(() => {
    let mounted = true;

    const loadMyProfileAndSelectManager: () => Promise<void> = async () => {
      try {
        const prof = await sp.profiles.myProperties() as unknown as {
          UserProfileProperties?: ProfileEntry[];
        };

        const props = new Map<string, string>(
          (prof?.UserProfileProperties ?? []).map((p: ProfileEntry) => [
            String(p?.Key ?? ""),
            String(p?.Value ?? ""),
          ])
        );

        const myDepartment = props.get("Department") || "";
        const myOffice = props.get("Office") || props.get("SPS-Location") || "";

        const chosen = pickDeptManager(myOffice, myDepartment);
        if (mounted && chosen) {
          setDeptMgrEmail(chosen.email);
          setDeptMgrName(chosen.name);
        }
      } catch (err) {
        console.warn("Could not load user profile / preselect manager:", err);
      }
    };

    loadMyProfileAndSelectManager().catch((err) => {
      console.warn("profile load failed:", err);
    });
    return () => { mounted = false; };
  }, [sp]);

  // ---------- Handlers ----------
  const onAddLine = (): void => {
    if (lineItems.length >= MAX_LINES) return;
    setLineItems([...lineItems, { description: "", qty: "", unitCost: "" }]);
  };

  const onRemoveLine = (idx: number): void => {
    setLineItems(lineItems.filter((_, i) => i !== idx));
  };

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    setFiles(e.target.files);
  };

  const addSelectedPreviousItems = (): void => {
    if (!selectedPrevKeys.length) return;
    const additions: ILineItem[] = [];
    for (const key of selectedPrevKeys) {
      const opt = prevItemOptions.find((o) => o.key === key);
      const data = opt?.data as { itemName: string; unitPrice: number; costCenter?: string } | undefined;
      if (!data) continue;
      additions.push({
        description: data.itemName,
        qty: "1",
        unitCost: String(data.unitPrice),
        costCenter: data.costCenter || undefined,
      });
    }
    if (!additions.length) return;

    const combined = [...lineItems, ...additions];
    setLineItems(combined);
    setSelectedPrevKeys([]);
  };

  const validate = (): string[] => {
    const errs: string[] = [];
    if (!expenseCategory) errs.push("Expense Category is required.");
    if (!vendorKey && !addingVendor) errs.push("Select a vendor or add a new one.");
    if (addingVendor && !newVendorName.trim()) errs.push("New vendor name is required.");
    if (!requestDate) errs.push("Request date is required.");
    if (!deptMgrEmail) errs.push("Department Manager is required.");
    if (!paymentMethod) errs.push("Select a payment method.");

    const anyValidLine = lineItems.some(
      (li) => li.description.trim() && toNum(li.qty) > 0 && toNum(li.unitCost) >= 0
    );
    if (!anyValidLine) errs.push("Enter at least one valid line item (description, qty > 0, unit cost).");

    if (paymentMethod === "Check Request") {
      if (!jobNo.trim()) errs.push("Job # is required for Check Request.");
      if (!taxId.trim()) errs.push("Tax ID# is required for Check Request.");
      if (!telephone.trim()) errs.push("Telephone # is required for Check Request.");
      if (toNum(checkAmount) <= 0) errs.push("Check Amount must be > 0.");
      if (!dateNeeded) errs.push("Date Needed is required.");
      if (!reason.trim()) errs.push("Reason is required.");
      if (!returnMethod) errs.push("Select Mail or Return to Requestor.");
    }
    return errs;
  };



  const ensureVendor = async (): Promise<number> => {
     if (!addingVendor) {
      if (vendorKey === undefined || vendorKey === null) {
         throw new Error("Vendor is required.");
       }
       return Number(vendorKey);
     }

     const title = newVendorName.trim();
     if (!title) {
       throw new Error("New vendor name is required.");
     }

     const payload: Record<string, unknown> = { Title: title };
     if (newVendorAddr.trim()) payload.Address = newVendorAddr.trim();
     if (newVendorWebsite.trim()) payload.Website = newVendorWebsite.trim();

     // Create the vendor
     const addResult = await sp.web.lists
       .getByTitle(LIST_NAMES.vendors)
       .items.add(payload) as unknown as PnpItemAddResult;

     // Try to get the ID directly from the add result
     const idFromAdd = getItemIdFromAddResult(addResult);

     let finalId: number | undefined;

     if (idFromAdd) {
       finalId = idFromAdd;
     } else {
       // Fallback: look up the vendor we just created
       const filters: string[] = [`Title eq '${spSingleQuote(title)}'`];
       if (newVendorAddr.trim()) {
         filters.push(`Address eq '${spSingleQuote(newVendorAddr.trim())}'`);
       }

       const candidates = await sp.web.lists
         .getByTitle(LIST_NAMES.vendors)
         .items.select("Id,Title,Address")
         .filter(filters.join(" and "))
         .orderBy("Id", false) // newest first
         .top(1)();

        finalId = candidates?.[0]?.Id;
     }

     if (!finalId) {
       console.error("Vendor addResult shape:", addResult);
       throw new Error("New vendor was created but its ID could not be determined.");
     }

     // Update dropdown options & select the new vendor
     setVendorOptions(prev => {
       const vendorOption: IVendorOption = {
         key: finalId,
         text: newVendorAddr.trim()
           ? `${title} — ${newVendorAddr.trim()}`
           : title,
         data: {
           id: finalId,
           address: newVendorAddr,
           website: newVendorWebsite,
         },
       };
       return [...prev, vendorOption];
     });

     setVendorKey(finalId);

     return finalId;
   };

  const uploadFiles = async (headerId: number): Promise<void> => {
    if (!headerId || !files || files.length === 0) return;

    const item = sp.web.lists.getByTitle(LIST_NAMES.requests).items.getById(headerId);

    for (let i = 0; i < files.length; i++) {
      const file = files.item(i)!;
      await item.attachmentFiles.add(file.name, file);
    }
  };

  const clearAllFields = (): void => {
    setExpenseCategory(undefined);
    setVendorKey(undefined);
    setAddingVendor(false);
    setNewVendorName("");
    setNewVendorAddr("");
    setNewVendorWebsite("");

    setDeptMgrEmail(undefined);
    setDeptMgrName(undefined);

    setRequestDate(undefined);
    setSalesTax("");
    setPaymentMethod(undefined);

    setLineItems([{ description: "", qty: "", unitCost: "" }]);

    setFiles(null);
    if (fileInputRef.current) fileInputRef.current.value = "";

    setJobNo("");
    setTaxId("");
    setTelephone("");
    setCheckAmount("");
    setDateNeeded(undefined);
    setReason("");
    setReturnMethod(undefined);

    setSelectedPrevKeys([]);
    setPrevItemOptions([]);
  };

  const onSubmit = async (): Promise<void> => {
    setMessage(null);
    const errs = validate();
    if (errs.length) {
      setMessage({ type: MessageBarType.error, text: errs.join(" \n") });
      return;
    }
    setBusy(true);
    try {
      const vendorId = await ensureVendor();

      const lineItemsForSave = lineItems
        .filter((li) => li.description.trim() && toNum(li.qty) > 0)
        .map((li) => ({
          ItemDescription: li.description.trim(),
          Quantity: toNum(li.qty),
          UnitCost: toNum(li.unitCost),
          CostCenter: li.costCenter || "",
        }));

      const headerPayload: Record<string, unknown> = {
        Title: `Purchase Request - ${new Date().toISOString()}`,
        ExpenseCategory: expenseCategory,
        RequestDate: requestDate?.toISOString(),
        Over200: isOver200,
        SalesTax: toNum(salesTax),
        PaymentMethod: paymentMethod,
        SubTotal: subTotal,
        Total: total,
        [REQUESTS_LINEITEMS_JSON_FIELD]: JSON.stringify(lineItemsForSave),
        DeptManagerEmail: deptMgrEmail || "",
        DeptManagerName: deptMgrName || "",
      };
      (headerPayload as Record<string, unknown>)[`${REQUESTS_VENDOR_LOOKUP_INTERNAL_NAME}Id`] = vendorId;

      if (paymentMethod === "Check Request") {
        headerPayload.JobNumber = jobNo;
        headerPayload.TaxID = taxId;
        headerPayload.Telephone = telephone;
        headerPayload.CheckAmount = toNum(checkAmount);
        headerPayload.DateNeeded = dateNeeded?.toISOString();
        headerPayload.Reason = reason;
        headerPayload.ReturnMethod = returnMethod;
      }

      // ===== CREATE HEADER =====
      const addResult = (await sp.web.lists.getByTitle(LIST_NAMES.requests).items.add(
        headerPayload
      )) as unknown as PnpItemAddResult;

      let headerId: number | undefined = getItemIdFromAddResult(addResult);

      // Fallback: query by exact Title if Id wasn't returned
      if (!headerId) {
        const title = String(headerPayload.Title || "");
        const candidates = await sp.web.lists
          .getByTitle(LIST_NAMES.requests)
          .items.select("Id,Title,Created")
          .filter(`Title eq '${spSingleQuote(title)}'`)
          .orderBy("Id", false)
          .top(1)();

        headerId = candidates?.[0]?.Id;
      }

      if (!headerId) {
        throw new Error("New request was created but its ID could not be determined.");
      }

      // ===== DETAIL LINES (optional if you rely solely on LineItemsJson) =====
      const lineList = sp.web.lists.getByTitle(LIST_NAMES.lines);
      for (const li of lineItems) {
        if (!(li.description.trim() && toNum(li.qty) > 0)) continue;
        const qty = toNum(li.qty);
        const price = toNum(li.unitCost);
        const lineTotal = qty * price;
        await lineList.items.add({
          Title: li.description,
          PO_IDId: headerId,
          ItemName: li.description,
          Qty: qty,
          UnitPrice: price,
          LineTotal: lineTotal,
          CostCenter: li.costCenter || null,
        });
      }

      // Attachments
      await uploadFiles(headerId);

      setMessage({ type: MessageBarType.success, text: "Request submitted successfully." });
      setTimeout(() => {
        try {
          clearAllFields();
        } catch (err) {
          console.warn("clearAllFields failed:", err);
        }
        window.location.assign("https://taylorwiseman.sharepoint.com/sites/PurchaseOrders/SitePages/Thank-you!.aspx");
      }, 600);
    } catch (e) {
      setMessage({ type: MessageBarType.error, text: `Submit failed: ${messageFromUnknownError(e)}` });
    } finally {
      setBusy(false);
    }
  };

  const handleSubmitClick = (): void => {
    onSubmit().catch((e) =>
      setMessage({ type: MessageBarType.error, text: `Submit failed: ${messageFromUnknownError(e)}` })
    );
  };

  const onRootKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === "Enter") {
      const tag = (e.target as HTMLElement)?.tagName?.toLowerCase();
      if (tag !== "textarea") {
        e.preventDefault();
        e.stopPropagation();
      }
    }
  };

  // ---------- UI ----------
  return (
    <div className="max-w-5xl mx-auto p-6" onKeyDown={onRootKeyDown}>
      <h1 className="text-2xl font-semibold mb-2">Purchase Request</h1>
      <p className="text-sm" style={{ color: "#6b7280", marginBottom: 16 }}>
        Fill in the details below. Fields with * are required.
      </p>

      {busy && (
        <div className="my-2">
          <Spinner size={SpinnerSize.medium} label="Working..." />
        </div>
      )}

      <Stack tokens={{ childrenGap: 12 }}>
        {/* Department Manager */}
        <Dropdown
          label="Department Manager *"
          placeholder="Select a manager"
          options={DEPT_MANAGER_OPTIONS}
          selectedKey={deptMgrEmail}
          onChange={(_, o): void => {
            const sel = DEPT_MANAGER_MAP.find(r => r.email === o?.key);
            setDeptMgrEmail(String(o?.key || ""));
            setDeptMgrName(sel?.name || "");
          }}
          required
        />

        <Dropdown
          label="Expense Category *"
          placeholder="Select a category"
          options={EXPENSE_CATEGORIES}
          selectedKey={expenseCategory}
          onChange={(_, o): void => setExpenseCategory(o?.key as string)}
          required
        />

        <Label required>Vendor</Label>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <div style={{ flex: 1 }}>
            <ComboBox
              allowFreeform={false}
              autoComplete="on"
              placeholder="Search/select vendor"
              options={vendorOptions}
              selectedKey={vendorKey}
              onChange={(_, option): void => {
                setVendorKey(option?.key);
                setPrevItemOptions([]);
                setSelectedPrevKeys([]);
              }}
              useComboBoxAsMenuWidth
            />
          </div>
          <DefaultButton
            type="button"
            text={addingVendor ? "Cancel" : "Add vendor"}
            onClick={(): void => setAddingVendor((v) => !v)}
          />
        </Stack>

        {!addingVendor && selectedVendor?.data?.address && (
          <TextField label="Vendor address" value={selectedVendor.data.address} readOnly />
        )}

        {addingVendor && (
          <div className="rounded-xl" style={{ border: "1px solid #e5e7eb", padding: 16, background: "#f9fafb" }}>
            <TextField
              label="Vendor name *"
              value={newVendorName}
              onChange={(_, v): void => setNewVendorName(v || "")}
              required
            />
            <TextField label="Vendor address" value={newVendorAddr} onChange={(_, v): void => setNewVendorAddr(v || "")} />
            <TextField
              label="Vendor website"
              value={newVendorWebsite}
              onChange={(_, v): void => setNewVendorWebsite(v || "")}
            />
          </div>
        )}

        <DatePicker
          label="Request date *"
          firstDayOfWeek={DayOfWeek.Sunday}
          value={requestDate}
          onSelectDate={(d): void => setRequestDate(d || undefined)}
        />

        <Separator>Reuse items previously purchased from this vendor</Separator>
        <Stack tokens={{ childrenGap: 8 }}>
          <ComboBox
            label="Pick previous items (multi-select)"
            placeholder={
              vendorKey
                ? prevItemOptions.length
                  ? "Search and select items..."
                  : "No previous items for this vendor"
                : "Select a vendor first"
            }
            multiSelect
            options={prevItemOptions}
            selectedKey={undefined}
            onChange={(_, option): void => {
              if (!option) return;
              setSelectedPrevKeys((prev) => {
                if (option.selected) return [...prev, option.key];
                return prev.filter((k) => k !== option.key);
              });
            }}
            useComboBoxAsMenuWidth
            allowFreeform={false}
            autoComplete="on"
            disabled={!vendorKey || !prevItemOptions.length}
          />
          <DefaultButton
            type="button"
            text="Add selected previous items"
            onClick={addSelectedPreviousItems}
            disabled={!selectedPrevKeys.length}
          />
        </Stack>

        <Separator>Line Items</Separator>
        <Stack tokens={{ childrenGap: 8 }}>
          {lineItems.map((li, idx) => (
            <div key={idx} className="rounded-2xl" style={{ border: "1px solid #e5e7eb", padding: 16 }}>
              <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                <TextField
                  label="Item Description *"
                  value={li.description}
                  onChange={(_, v): void => {
                    const copy = [...lineItems];
                    copy[idx].description = v || "";
                    setLineItems(copy);
                  }}
                  styles={{ root: { flex: 2 } }}
                  required
                />
                <TextField
                  label="Quantity *"
                  type="text"
                  inputMode="decimal"
                  value={li.qty}
                  onChange={(_, v): void => {
                    const next = v ?? "";
                    if (!isMoneyLike(next)) return;
                    const copy = [...lineItems];
                    copy[idx].qty = next;
                    setLineItems(copy);
                  }}
                  styles={{ root: { width: 160 } }}
                  required
                />
                <TextField
                  label="Unit Cost *"
                  type="text"
                  inputMode="decimal"
                  value={li.unitCost}
                  onChange={(_, v): void => {
                    const next = v ?? "";
                    if (!isMoneyLike(next)) return;
                    const copy = [...lineItems];
                    copy[idx].unitCost = next;
                    setLineItems(copy);
                  }}
                  styles={{ root: { width: 180 } }}
                  required
                />
                <Dropdown
                  label="Cost Center"
                  options={COST_CENTERS}
                  selectedKey={li.costCenter}
                  onChange={(_, o): void => {
                    const copy = [...lineItems];
                    copy[idx].costCenter = (o?.key as string) || undefined;
                    setLineItems(copy);
                  }}
                  styles={{ root: { width: 240 } }}
                />
                <div className="self-end" style={{ marginLeft: "auto" }}>
                  <IconButton
                    type="button"
                    iconProps={{ iconName: "Delete" }}
                    aria-label="Remove line"
                    onClick={(): void => onRemoveLine(idx)}
                    disabled={lineItems.length === 1}
                  />
                </div>
              </Stack>
              <div className="text-sm" style={{ color: "#4b5563", marginTop: 8 }}>
                Line total: <strong>{currency(toNum(li.qty) * toNum(li.unitCost))}</strong>
              </div>
            </div>
          ))}
          <div>
            <DefaultButton type="button" text="Add line" onClick={onAddLine} disabled={lineItems.length >= MAX_LINES} />
          </div>
        </Stack>

        <div>
          <Label>Supporting documentation (quotes, etc.)</Label>
          <input ref={fileInputRef} type="file" multiple onChange={onFileChange} />
        </div>

        <TextField
          label="Sales Tax"
          type="text"
          inputMode="decimal"
          value={salesTax}
          onChange={(_, v): void => {
            const next = v ?? "";
            if (!isMoneyLike(next)) return;
            setSalesTax(next);
          }}
          styles={{ root: { maxWidth: 240 } }}
          prefix="$"
          placeholder="0.00"
        />

        <ChoiceGroup
          label="Payment Method *"
          options={PAYMENT_METHODS}
          selectedKey={paymentMethod}
          onChange={(_, o): void => setPaymentMethod(o?.key)}
        />

        {paymentMethod === "Check Request" && (
          <div className="rounded-2xl" style={{ border: "1px solid #e5e7eb", padding: 16, background: "#f9fafb" }}>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
              <TextField
                label="Job # *"
                value={jobNo}
                onChange={(_, v): void => setJobNo(v || "")}
                required
                styles={{ root: { width: 200 } }}
              />
              <TextField
                label="Tax ID# *"
                value={taxId}
                onChange={(_, v): void => setTaxId(v || "")}
                required
                styles={{ root: { width: 220 } }}
              />
              <TextField
                label="Telephone # *"
                value={telephone}
                onChange={(_, v): void => setTelephone(v || "")}
                required
                styles={{ root: { width: 220 } }}
              />
              <TextField
                label="Check Amount *"
                type="text"
                inputMode="decimal"
                value={checkAmount}
                onChange={(_, v): void => {
                  const next = v ?? "";
                  if (!isMoneyLike(next)) return;
                  setCheckAmount(next);
                }}
                required
                styles={{ root: { width: 200 } }}
              />
              <DatePicker label="Date Needed *" value={dateNeeded} onSelectDate={(d): void => setDateNeeded(d || undefined)} />
              <Dropdown
                label="Return Method *"
                options={[
                  { key: "Mail", text: "Mail" },
                  { key: "Return to Requestor", text: "Return to Requestor" },
                ]}
                selectedKey={returnMethod}
                onChange={(_, o): void => setReturnMethod((o?.key as ReturnMethod) || undefined)}
                styles={{ root: { width: 240 } }}
              />
            </Stack>
            <TextField label="Reason *" multiline autoAdjustHeight value={reason} onChange={(_, v): void => setReason(v || "")} required />
          </div>
        )}

        <div className="rounded-2xl" style={{ border: "1px solid #e5e7eb", padding: 16 }}>
          <div className="text-sm" style={{ color: "#4b5563" }}>
            Subtotal: <strong>{currency(subTotal)}</strong>
          </div>
          <div className="text-sm" style={{ color: "#4b5563" }}>
            Sales Tax: <strong>{currency(toNum(salesTax))}</strong>
          </div>
          <Separator />
          <div className="text-base">
            Total: <strong>{currency(total)}</strong>
          </div>
        </div>

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="Submit Request"
            type="button"
            disabled={busy}
            onClick={(e): void => { e.preventDefault(); e.stopPropagation(); handleSubmitClick(); }}
          />
        </Stack>

        {message && (
          <div className="mt-2">
            <MessageBar
              messageBarType={message.type}
              isMultiline={true}
              onDismiss={(): void => setMessage(null)}
            >
              {message.text}
            </MessageBar>
          </div>
        )}
      </Stack>
    </div>
  );
};

export default PurchaseRequestForm;
