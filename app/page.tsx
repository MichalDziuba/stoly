/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react-hooks/set-state-in-effect */
/* eslint-disable @typescript-eslint/no-explicit-any */
"use client";

import React, { useState, useEffect, useRef, ChangeEvent } from "react";

export interface Guest {
  id: string;
  fullName: string;
  tableId: string | null;
  seatIndex?: number | null;
}

export interface Table {
  id: string;
  name: string;
  shape: "round" | "rectangle" | "rectangle-one-side";
  capacity: number;
  x: number;
  y: number;
  rotation: number;
}

// Ikony SVG
const IconUpload = () => (
  <svg
    className="w-4 h-4"
    fill="none"
    stroke="currentColor"
    viewBox="0 0 24 24"
  >
    <path
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"
    />
  </svg>
);

const IconSave = () => (
  <svg
    className="w-4 h-4"
    fill="none"
    stroke="currentColor"
    viewBox="0 0 24 24"
  >
    <path
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"
    />
  </svg>
);

const IconPlus = () => (
  <svg
    className="w-4 h-4"
    fill="none"
    stroke="currentColor"
    viewBox="0 0 24 24"
  >
    <path
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      d="M12 4v16m8-8H4"
    />
  </svg>
);

const IconTrash = () => (
  <svg
    className="w-4 h-4"
    fill="none"
    stroke="currentColor"
    viewBox="0 0 24 24"
  >
    <path
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"
    />
  </svg>
);

const IconMaximize = () => (
  <svg
    className="w-4 h-4"
    fill="none"
    stroke="currentColor"
    viewBox="0 0 24 24"
  >
    <path
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      d="M4 8V4m0 0h4M4 4l5 5m11-1V4m0 0h-4m4 0l-5 5M4 16v4m0 0h4m-4 0l5-5m11 5l-5-5m5 5v-4m0 4h-4"
    />
  </svg>
);

const IconClose = () => (
  <svg
    className="w-5 h-5"
    fill="none"
    stroke="currentColor"
    viewBox="0 0 24 24"
  >
    <path
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      d="M6 18L18 6M6 6l12 12"
    />
  </svg>
);

const getInitials = (name: string): string => {
  if (!name) return "?";
  const parts = name.trim().split(/\s+/);
  if (parts.length >= 2) {
    return (parts[0][0] + parts[1][0]).toUpperCase();
  }
  return name.trim().substring(0, 2).toUpperCase();
};

const normalizeGuestsAndTables = (
  guestsList: Guest[],
  tablesList: Table[],
): Guest[] => {
  const tableSeatsMap: Record<string, Set<number>> = {};
  tablesList.forEach((t) => {
    tableSeatsMap[t.id] = new Set<number>();
  });

  guestsList.forEach((g) => {
    if (
      g.tableId &&
      tableSeatsMap[g.tableId] &&
      typeof g.seatIndex === "number" &&
      g.seatIndex >= 0
    ) {
      tableSeatsMap[g.tableId].add(g.seatIndex);
    }
  });

  return guestsList.map((g) => {
    if (!g.tableId || !tableSeatsMap[g.tableId]) {
      return { ...g, tableId: g.tableId || null, seatIndex: null };
    }

    if (typeof g.seatIndex === "number" && g.seatIndex >= 0) {
      return g;
    }

    const table = tablesList.find((t) => t.id === g.tableId);
    const capacity = table ? table.capacity : 100;
    const occupied = tableSeatsMap[g.tableId];

    let freeSeat = 0;
    while (occupied.has(freeSeat) && freeSeat < capacity) {
      freeSeat++;
    }

    occupied.add(freeSeat);
    return { ...g, seatIndex: freeSeat };
  });
};

const loadXLSX = async () => {
  if (typeof window !== "undefined" && (window as any).XLSX) {
    return (window as any).XLSX;
  }
  try {
    const XLSX = await import("xlsx");
    return XLSX.default || XLSX;
  } catch (e) {
    return new Promise((resolve, reject) => {
      if (typeof window === "undefined") return reject(e);
      if ((window as any).XLSX) return resolve((window as any).XLSX);
      const script = document.createElement("script");
      script.src =
        "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
      script.onload = () => resolve((window as any).XLSX);
      script.onerror = () =>
        reject(new Error("Nie udało się załadować biblioteki XLSX"));
      document.head.appendChild(script);
    });
  }
};

export default function App() {
  const [hasMounted, setHasMounted] = useState<boolean>(false);

  const [guests, setGuests] = useState<Guest[]>([]);
  const [tables, setTables] = useState<Table[]>([]);

  const [selectedTableId, setSelectedTableId] = useState<string | null>(null);
  const [isDetailModalOpen, setIsDetailModalOpen] = useState<boolean>(false);

  const [draggedGuestId, setDraggedGuestId] = useState<string | null>(null);
  const [dragOverTarget, setDragOverTarget] = useState<string | null>(null);
  const [selectedSeatForSwap, setSelectedSeatForSwap] = useState<number | null>(
    null,
  );
  const [detailSearchQuery, setDetailSearchQuery] = useState<string>("");

  const [newTableName, setNewTableName] = useState<string>("");
  const [newTableShape, setNewTableShape] = useState<
    "round" | "rectangle" | "rectangle-one-side"
  >("round");
  const [newTableCapacity, setNewTableCapacity] = useState<number>(8);

  const [quickGuestName, setQuickGuestName] = useState<string>("");

  const [zoom, setZoom] = useState<number>(1);
  const [alwaysShowNames, setAlwaysShowNames] = useState<boolean>(false);

  const [isDraggingTable, setIsDraggingTable] = useState<boolean>(false);
  const [activeDragTableId, setActiveDragTableId] = useState<string | null>(
    null,
  );
  const dragStartPosRef = useRef<{
    clientX: number;
    clientY: number;
    tableX: number;
    tableY: number;
  } | null>(null);

  const [statusMessage, setStatusMessage] = useState<string>(
    "Brak wczytanego planu",
  );
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const canvasRef = useRef<HTMLDivElement>(null);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 3200);
  };

  const saveToLocalMemory = (tablesData: Table[], guestsData: Guest[]) => {
    if (typeof window !== "undefined") {
      try {
        localStorage.setItem("wedding_tables", JSON.stringify(tablesData));
        localStorage.setItem("wedding_guests", JSON.stringify(guestsData));
      } catch (e) {
        console.error("Błąd zapisu w localStorage:", e);
      }
    }
  };

  useEffect(() => {
    setHasMounted(true);
    if (typeof window !== "undefined") {
      const savedTables = localStorage.getItem("wedding_tables");
      const savedGuests = localStorage.getItem("wedding_guests");

      let loadedTables: Table[] = [];
      let loadedGuests: Guest[] = [];

      if (savedTables) {
        try {
          loadedTables = JSON.parse(savedTables);
        } catch (e) {}
      }
      if (savedGuests) {
        try {
          loadedGuests = JSON.parse(savedGuests);
        } catch (e) {}
      }

      if (loadedTables.length === 0 && loadedGuests.length === 0) {
        loadedTables = [
          {
            id: "stol-prezydencki",
            name: "Stół Prezydencki",
            shape: "rectangle",
            capacity: 6,
            x: 300,
            y: 80,
            rotation: 0,
          },
          {
            id: "stol-1",
            name: "Stół 1",
            shape: "round",
            capacity: 8,
            x: 150,
            y: 280,
            rotation: 0,
          },
          {
            id: "stol-2",
            name: "Stół 2",
            shape: "round",
            capacity: 8,
            x: 450,
            y: 280,
            rotation: 0,
          },
        ];
      }

      const normalizedGuests = normalizeGuestsAndTables(
        loadedGuests,
        loadedTables,
      );
      setTables(loadedTables);
      setGuests(normalizedGuests);

      if (loadedTables.length > 0 || loadedGuests.length > 0) {
        setStatusMessage(
          `Wczytano: ${loadedTables.length} stołów, ${loadedGuests.length} gości`,
        );
      } else {
        setStatusMessage("Pusty plan - dodaj stoły lub zaimportuj dane");
      }
    }
  }, []);

  const handleClearAll = () => {
    if (window.confirm("Czy na pewno chcesz usunąć wszystkie stoły i gości?")) {
      setTables([]);
      setGuests([]);
      setSelectedTableId(null);
      saveToLocalMemory([], []);
      setStatusMessage("Wyczyszczono plan");
      showToast("Wyczyszczono cały plan");
    }
  };

  const handleAddTable = () => {
    const tableName = newTableName.trim() || `Stół ${tables.length + 1}`;
    const newTable: Table = {
      id: `stol-${Date.now()}`,
      name: tableName,
      shape: newTableShape,
      capacity: Math.max(1, newTableCapacity),
      x: 100 + (tables.length % 4) * 200,
      y: 100 + Math.floor(tables.length / 4) * 180,
      rotation: 0,
    };

    const updatedTables = [...tables, newTable];
    setTables(updatedTables);
    saveToLocalMemory(updatedTables, guests);
    setNewTableName("");
    setSelectedTableId(newTable.id);
    showToast(`Dodano stół "${tableName}"`);
  };

  const handleRemoveTable = (tableId: string) => {
    const tableToRemove = tables.find((t) => t.id === tableId);
    if (!tableToRemove) return;

    if (window.confirm(`Czy na pewno chcesz usunąć "${tableToRemove.name}"?`)) {
      const updatedTables = tables.filter((t) => t.id !== tableId);
      const updatedGuests = guests.map((g) =>
        g.tableId === tableId ? { ...g, tableId: null, seatIndex: null } : g,
      );
      setTables(updatedTables);
      setGuests(updatedGuests);
      if (selectedTableId === tableId) setSelectedTableId(null);
      saveToLocalMemory(updatedTables, updatedGuests);
      showToast(`Usunięto stół "${tableToRemove.name}"`);
    }
  };

  const handleAssignGuestToTableSeat = (
    guestId: string,
    targetTableId: string | null,
    targetSeatIndex: number | null = null,
  ) => {
    const guestObj = guests.find((g) => g.id === guestId);
    if (!guestObj) return;

    if (targetTableId === null) {
      const updated = guests.map((g) =>
        g.id === guestId ? { ...g, tableId: null, seatIndex: null } : g,
      );
      const normalized = normalizeGuestsAndTables(updated, tables);
      setGuests(normalized);
      saveToLocalMemory(tables, normalized);
      showToast(`Odstawiono ${guestObj.fullName} do nieprzypisanych`);
      return;
    }

    const targetTable = tables.find((t) => t.id === targetTableId);
    if (!targetTable) return;

    let finalSeatIndex = targetSeatIndex;
    if (finalSeatIndex === null) {
      const occupiedSeats = new Set(
        guests
          .filter(
            (g) =>
              g.tableId === targetTableId &&
              g.id !== guestId &&
              typeof g.seatIndex === "number",
          )
          .map((g) => g.seatIndex as number),
      );
      for (let i = 0; i < targetTable.capacity; i++) {
        if (!occupiedSeats.has(i)) {
          finalSeatIndex = i;
          break;
        }
      }
      if (finalSeatIndex === null) {
        finalSeatIndex = 0;
      }
    }

    const occupantAtTarget = guests.find(
      (g) =>
        g.tableId === targetTableId &&
        g.seatIndex === finalSeatIndex &&
        g.id !== guestId,
    );

    const updated = guests.map((g) => {
      if (g.id === guestId) {
        return { ...g, tableId: targetTableId, seatIndex: finalSeatIndex };
      }
      if (occupantAtTarget && g.id === occupantAtTarget.id) {
        return {
          ...g,
          tableId: guestObj.tableId,
          seatIndex: guestObj.seatIndex,
        };
      }
      return g;
    });

    const normalized = normalizeGuestsAndTables(updated, tables);
    setGuests(normalized);
    saveToLocalMemory(tables, normalized);
    showToast(`Przeniesiono ${guestObj.fullName} ➔ ${targetTable.name}`);
  };

  const handleExportJSON = () => {
    const dataStr =
      "data:text/json;charset=utf-8," +
      encodeURIComponent(JSON.stringify({ tables, guests }, null, 2));
    const downloadAnchor = document.createElement("a");
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute(
      "download",
      `plan_wesela_${new Date().toISOString().slice(0, 10)}.json`,
    );
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
    showToast("Pobrano plik JSON!");
  };

  const handleImportJSON = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target?.result as string);
        if (Array.isArray(parsed.tables) && Array.isArray(parsed.guests)) {
          const normalizedGuests = normalizeGuestsAndTables(
            parsed.guests,
            parsed.tables,
          );
          setTables(parsed.tables);
          setGuests(normalizedGuests);
          saveToLocalMemory(parsed.tables, normalizedGuests);
          setStatusMessage(
            `Załadowano: ${parsed.tables.length} stołów, ${parsed.guests.length} osób`,
          );
          showToast("Załadowano układ z pliku JSON!");
        }
      } catch (err) {
        showToast("Błąd: Nieprawidłowy plik JSON.");
      }
    };
    reader.readAsText(file);
  };

  const handleExcelImport = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setStatusMessage("Wczytywanie pliku Excel...");
      const XLSX = await loadXLSX();
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
      });

      if (!rawRows || rawRows.length === 0) {
        showToast("Plik Excel jest pusty.");
        return;
      }

      const importedGuests: Guest[] = [];
      const detectedTablesMap: Record<string, number> = {};

      rawRows.forEach((row, rowIndex) => {
        if (!Array.isArray(row)) return;

        const nonArrayValues = row
          .map((cell) => String(cell).trim())
          .filter((cellVal) => cellVal !== "");

        if (nonArrayValues.length === 0) return;

        if (
          rowIndex === 0 &&
          nonArrayValues.some((v) =>
            /imie|imię|nazwisko|osoba|stół|stol|table/i.test(v),
          )
        ) {
          return;
        }

        let fullName = "";
        let tableName = "";

        if (nonArrayValues.length >= 2) {
          fullName = nonArrayValues[0];
          tableName = nonArrayValues[1];
        } else if (nonArrayValues.length === 1) {
          fullName = nonArrayValues[0];
        }

        if (fullName) {
          importedGuests.push({
            id: `guest-${rowIndex}-${Date.now()}`,
            fullName: fullName,
            tableId: tableName || null,
          });

          if (tableName) {
            detectedTablesMap[tableName] =
              (detectedTablesMap[tableName] || 0) + 1;
          }
        }
      });

      if (importedGuests.length === 0) {
        showToast("Nie znaleziono osób w pliku Excel.");
        return;
      }

      const newTablesList = [...tables];

      Object.entries(detectedTablesMap).forEach(([tableName, count]) => {
        const existingTable = newTablesList.find(
          (t) => t.name.toLowerCase().trim() === tableName.toLowerCase().trim(),
        );

        if (existingTable) {
          if (existingTable.capacity < count) {
            existingTable.capacity = count;
          }
        } else {
          const tableId = `stol-${tableName.toLowerCase().replace(/\s+/g, "-")}-${Date.now()}`;
          newTablesList.push({
            id: tableId,
            name: tableName,
            shape: count > 10 ? "rectangle" : "round",
            capacity: Math.max(count, 8),
            x: 100 + (newTablesList.length % 3) * 260,
            y: 100 + Math.floor(newTablesList.length / 3) * 220,
            rotation: 0,
          });
        }
      });

      const finalizedGuests = importedGuests.map((g) => {
        if (!g.tableId) return g;
        const matchedTable = newTablesList.find(
          (t) =>
            t.name.toLowerCase().trim() === g.tableId?.toLowerCase().trim(),
        );
        return {
          ...g,
          tableId: matchedTable ? matchedTable.id : null,
        };
      });

      const normalizedGuests = normalizeGuestsAndTables(
        finalizedGuests,
        newTablesList,
      );

      setTables(newTablesList);
      setGuests(normalizedGuests);

      saveToLocalMemory(newTablesList, normalizedGuests);

      setStatusMessage(`Zaimportowano ${normalizedGuests.length} osób ✓`);
      showToast(`Zaimportowano ${normalizedGuests.length} osób!`);
    } catch (err) {
      console.error("Błąd importu Excel:", err);
      showToast("Błąd odczytu pliku Excel.");
    }
  };

  const handleDragStartGuest = (e: React.DragEvent, guestId: string) => {
    e.stopPropagation();
    setDraggedGuestId(guestId);
    e.dataTransfer.setData("text/plain", guestId);
  };

  const handleDragOverTarget = (e: React.DragEvent, targetId: string) => {
    e.preventDefault();
    e.stopPropagation();
    if (dragOverTarget !== targetId) {
      setDragOverTarget(targetId);
    }
  };

  const handleDragLeaveTarget = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverTarget(null);
  };

  const handleDropOnTable = (e: React.DragEvent, tableId: string) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverTarget(null);
    const guestId = e.dataTransfer.getData("text/plain") || draggedGuestId;
    if (guestId) {
      handleAssignGuestToTableSeat(guestId, tableId, null);
    }
    setDraggedGuestId(null);
  };

  const handleDropOnSeat = (
    e: React.DragEvent,
    tableId: string,
    seatIndex: number,
  ) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverTarget(null);
    const guestId = e.dataTransfer.getData("text/plain") || draggedGuestId;
    if (guestId) {
      handleAssignGuestToTableSeat(guestId, tableId, seatIndex);
    }
    setDraggedGuestId(null);
  };

  const handleDropUnassigned = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverTarget(null);
    const guestId = e.dataTransfer.getData("text/plain") || draggedGuestId;
    if (guestId) {
      handleAssignGuestToTableSeat(guestId, null, null);
    }
    setDraggedGuestId(null);
  };

  const handleUpdateTableProps = (id: string, key: keyof Table, value: any) => {
    const updatedTables = tables.map((t) =>
      t.id === id ? { ...t, [key]: value } : t,
    );
    const normalizedGuests = normalizeGuestsAndTables(guests, updatedTables);
    setTables(updatedTables);
    setGuests(normalizedGuests);
    saveToLocalMemory(updatedTables, normalizedGuests);
  };

  const handleAddQuickGuestToSelectedTable = () => {
    if (!selectedTableId || !quickGuestName.trim()) return;

    const newGuest: Guest = {
      id: `guest-${Date.now()}`,
      fullName: quickGuestName.trim(),
      tableId: selectedTableId,
      seatIndex: null,
    };
    const updated = [...guests, newGuest];
    const normalizedGuests = normalizeGuestsAndTables(updated, tables);
    setGuests(normalizedGuests);
    saveToLocalMemory(tables, normalizedGuests);
    setQuickGuestName("");
    showToast(`Dodano "${newGuest.fullName}" do stołu`);
  };

  const handleMouseDownTable = (e: React.MouseEvent, table: Table) => {
    e.stopPropagation();
    dragStartPosRef.current = {
      clientX: e.clientX,
      clientY: e.clientY,
      tableX: table.x,
      tableY: table.y,
    };
    setActiveDragTableId(table.id);
    setIsDraggingTable(false);
  };

  const handleMouseMoveCanvas = (e: React.MouseEvent) => {
    if (!activeDragTableId || !dragStartPosRef.current) return;

    const deltaX = (e.clientX - dragStartPosRef.current.clientX) / zoom;
    const deltaY = (e.clientY - dragStartPosRef.current.clientY) / zoom;

    if (!isDraggingTable && (Math.abs(deltaX) > 4 || Math.abs(deltaY) > 4)) {
      setIsDraggingTable(true);
    }

    if (isDraggingTable || Math.abs(deltaX) > 4 || Math.abs(deltaY) > 4) {
      const newX = Math.max(10, dragStartPosRef.current.tableX + deltaX);
      const newY = Math.max(10, dragStartPosRef.current.tableY + deltaY);

      setTables((prev) =>
        prev.map((t) =>
          t.id === activeDragTableId ? { ...t, x: newX, y: newY } : t,
        ),
      );
    }
  };

  const handleMouseUpCanvas = () => {
    if (activeDragTableId) {
      if (!isDraggingTable) {
        const clickedTable = tables.find((t) => t.id === activeDragTableId);
        if (clickedTable) {
          setSelectedTableId(clickedTable.id);
        }
      } else {
        saveToLocalMemory(tables, guests);
      }
    }
    setActiveDragTableId(null);
    setIsDraggingTable(false);
    dragStartPosRef.current = null;
  };

  const handleWheelCanvas = (e: React.WheelEvent) => {
    if (e.ctrlKey || e.metaKey || e.altKey) {
      e.preventDefault();
      const delta = e.deltaY < 0 ? 0.1 : -0.1;
      setZoom((prev) =>
        Math.min(2.5, Math.max(0.5, parseFloat((prev + delta).toFixed(2)))),
      );
    }
  };

  if (!hasMounted) {
    return (
      <div className="flex h-screen w-screen bg-slate-900 items-center justify-center text-slate-300 font-sans">
        Ładowanie planera...
      </div>
    );
  }

  const activeTable = tables.find((t) => t.id === selectedTableId);

  const activeTableGuestsMap: Record<number, Guest> = {};
  if (activeTable) {
    guests
      .filter((g) => g.tableId === activeTable.id)
      .forEach((g) => {
        if (typeof g.seatIndex === "number" && g.seatIndex >= 0) {
          activeTableGuestsMap[g.seatIndex] = g;
        }
      });
  }

  return (
    <div className="flex h-screen w-screen bg-slate-900 text-slate-100 font-sans overflow-hidden relative select-none">
      {toastMessage && (
        <div className="fixed top-4 right-4 z-50 bg-amber-500 text-slate-950 px-4 py-2 rounded-lg font-bold shadow-2xl animate-bounce text-xs">
          {toastMessage}
        </div>
      )}

      {/* --- LEWY PANEL SIDEBAR --- */}
      <div className="w-80 lg:w-96 bg-slate-800 border-r border-slate-700 flex flex-col h-full z-10 shadow-2xl shrink-0">
        <div className="p-4 border-b border-slate-700 bg-slate-850">
          <div className="flex justify-between items-center mb-1">
            <h1 className="text-lg font-bold text-amber-400 flex items-center gap-2">
              🏰 Planer Stołów Weselnych
            </h1>
            <button
              onClick={handleClearAll}
              className="text-[10px] bg-red-950/80 text-red-300 hover:bg-red-900 border border-red-700/60 px-2 py-0.5 rounded transition"
              title="Wyczyść cały plan do zera"
            >
              Wyczyść
            </button>
          </div>
          <p className="text-xs text-slate-400">{statusMessage}</p>

          <div className="mt-3 space-y-2">
            <label className="flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white py-2 px-3 rounded-md cursor-pointer text-xs font-semibold transition shadow">
              <IconUpload />
              <span>Importuj z pliku Excel (.xlsx)</span>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleExcelImport}
                className="hidden"
              />
            </label>

            <div className="flex gap-2">
              <button
                onClick={handleExportJSON}
                className="flex-1 flex items-center justify-center gap-1 bg-slate-700 hover:bg-slate-600 text-slate-200 py-1.5 px-2 rounded text-xs border border-slate-600 transition"
              >
                <IconSave /> Pobierz JSON
              </button>
              <label className="flex-1 flex items-center justify-center gap-1 bg-slate-700 hover:bg-slate-600 text-slate-200 py-1.5 px-2 rounded text-xs border border-slate-600 transition cursor-pointer">
                <span>Wczytaj JSON</span>
                <input
                  type="file"
                  accept=".json"
                  onChange={handleImportJSON}
                  className="hidden"
                />
              </label>
            </div>
          </div>
        </div>

        {/* FORMULARZ NOWEGO STOŁU */}
        <div className="p-4 border-b border-slate-700 bg-slate-800/80">
          <h2 className="text-xs font-semibold text-slate-300 uppercase tracking-wider mb-2">
            Dodaj Nowy Stół
          </h2>
          <div className="space-y-2">
            <input
              type="text"
              placeholder="Nazwa (np. Stół Rodziców)"
              value={newTableName}
              onChange={(e) => setNewTableName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleAddTable()}
              className="w-full bg-slate-900 border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-indigo-500"
            />
            <div className="flex gap-2">
              <select
                value={newTableShape}
                onChange={(e) => setNewTableShape(e.target.value as any)}
                className="bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white flex-1 focus:outline-none"
              >
                <option value="round">Okrągły 🟣</option>
                <option value="rectangle">Prostokątny (2 strony) 🟦</option>
                <option value="rectangle-one-side">
                  Prostokątny (1 strona) ▭
                </option>
              </select>
              <input
                type="number"
                min="1"
                max="50"
                value={newTableCapacity}
                onChange={(e) =>
                  setNewTableCapacity(parseInt(e.target.value) || 8)
                }
                className="w-16 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white text-center focus:outline-none"
                title="Liczba krzesełek"
              />
              <button
                onClick={handleAddTable}
                className="bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-1 rounded text-xs font-semibold flex items-center gap-1 transition shrink-0"
              >
                <IconPlus /> Dodaj
              </button>
            </div>
          </div>
        </div>

        {/* LISTA NIEPRZYPISANYCH GOŚCI */}
        <div
          onDragOver={(e) => handleDragOverTarget(e, "unassigned")}
          onDragLeave={handleDragLeaveTarget}
          onDrop={handleDropUnassigned}
          className={`flex-1 overflow-y-auto p-4 space-y-2 transition-colors ${
            dragOverTarget === "unassigned"
              ? "bg-amber-950/40 border-2 border-dashed border-amber-400/80"
              : ""
          }`}
        >
          <div className="flex justify-between items-center mb-2">
            <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider">
              Nieprzypisani Goście ({guests.filter((g) => !g.tableId).length})
            </span>
            <span className="text-[10px] text-slate-500">
              Przeciągnij osobę na stół
            </span>
          </div>

          {guests.filter((g) => !g.tableId).length === 0 ? (
            <div className="text-center py-8 text-slate-500 text-xs italic border border-dashed border-slate-700/60 rounded-lg p-4">
              {guests.length === 0
                ? "Brak wprowadzonych gości. Wczytaj plik Excel lub dodaj stół i wpisz gości."
                : "Wszyscy goście mają przypisane stoły!"}
            </div>
          ) : (
            guests
              .filter((g) => !g.tableId)
              .map((guest) => (
                <div
                  key={guest.id}
                  draggable
                  onDragStart={(e) => handleDragStartGuest(e, guest.id)}
                  className="bg-slate-700/80 border border-slate-600 rounded-lg p-2.5 flex justify-between items-center hover:bg-slate-700 transition cursor-grab active:cursor-grabbing shadow-sm"
                >
                  <span
                    className="text-xs text-slate-100 font-semibold truncate pr-2 flex items-center gap-1.5"
                    title={guest?.fullName|| guest}
                  >
                    <span className="text-amber-400">✋</span> {guest.fullName}
                  </span>
                  {tables.length > 0 && (
                    <select
                      onChange={(e) => {
                        if (e.target.value) {
                          handleAssignGuestToTableSeat(
                            guest.id,
                            e.target.value,
                            null,
                          );
                        }
                      }}
                      value=""
                      className="bg-slate-800 text-[11px] text-amber-300 border border-slate-600 rounded px-1.5 py-0.5 focus:outline-none shrink-0 cursor-pointer"
                    >
                      <option value="" disabled>
                        Usadź przy...
                      </option>
                      {tables.map((t) => (
                        <option key={t.id} value={t.id}>
                          {t.name}
                        </option>
                      ))}
                    </select>
                  )}
                </div>
              ))
          )}
        </div>
      </div>

      {/* --- ŚRODOK: INTERAKTYWNA SALA --- */}
      <div className="flex-1 flex flex-col h-full bg-slate-950 relative overflow-hidden">
        {/* PASEK NARZĘDZI SALI */}
        <div className="min-h-12 bg-slate-900 border-b border-slate-800 px-3 py-1.5 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-400 z-20">
          <div className="flex items-center gap-2 min-w-0">
            <span className="bg-slate-800 text-amber-300 border border-amber-500/30 px-2.5 py-1 rounded-md text-[11px] font-medium flex items-center gap-1.5 truncate shadow-sm">
              💡 <strong>Instrukcja:</strong> Przeciągaj gości myszką na dowolny
              stół lub kliknij stół, aby go edytować.
            </span>
          </div>

          <div className="flex items-center gap-2 bg-slate-800/90 p-1 rounded-lg border border-slate-700 flex-wrap shrink-0 shadow-inner">
            <button
              onClick={() => {
                setZoom(1);
                setSelectedTableId(null);
              }}
              className="bg-indigo-600 hover:bg-indigo-500 text-white text-[11px] font-semibold px-2.5 py-1 rounded transition flex items-center gap-1 shadow-sm shrink-0"
            >
              🔍 Widok domyślny
            </button>

            <div className="flex items-center gap-1 bg-slate-900/60 px-1.5 py-0.5 rounded border border-slate-700/50">
              <button
                onClick={() =>
                  setZoom((z) =>
                    Math.max(0.5, parseFloat((z - 0.2).toFixed(2))),
                  )
                }
                className="w-5 h-5 bg-slate-700 hover:bg-slate-600 rounded text-slate-200 font-bold flex items-center justify-center transition shrink-0 text-xs"
              >
                -
              </button>
              <span className="font-mono text-amber-400 font-bold min-w-9 text-center text-[11px]">
                {Math.round(zoom * 100)}%
              </span>
              <button
                onClick={() =>
                  setZoom((z) =>
                    Math.min(2.5, parseFloat((z + 0.2).toFixed(2))),
                  )
                }
                className="w-5 h-5 bg-slate-700 hover:bg-slate-600 rounded text-slate-200 font-bold flex items-center justify-center transition shrink-0 text-xs"
              >
                +
              </button>
            </div>

            <label className="flex items-center gap-1 cursor-pointer text-[11px] text-slate-300 px-1 hover:text-white transition">
              <input
                type="checkbox"
                checked={alwaysShowNames}
                onChange={(e) => setAlwaysShowNames(e.target.checked)}
                className="rounded accent-indigo-500 cursor-pointer"
              />
              <span>Pokaż nazwiska</span>
            </label>
          </div>

          <div className="text-slate-400 text-[11px] shrink-0 bg-slate-800/80 px-2.5 py-1 rounded-md border border-slate-700/80 font-medium">
            Wszyscy goście:{" "}
            <strong className="text-amber-400 font-bold">
              {guests.length}
            </strong>
          </div>
        </div>

        {/* INTERAKTYWNY CANVAS SALI */}
        <div
          ref={canvasRef}
          onMouseMove={handleMouseMoveCanvas}
          onMouseUp={handleMouseUpCanvas}
          onWheel={handleWheelCanvas}
          onClick={() => setSelectedTableId(null)}
          className="flex-1 relative overflow-auto bg-[radial-gradient(#334155_1px,transparent_1px)] bg-size-[24px_24px] select-none scroll-smooth cursor-crosshair"
        >
          {tables.length === 0 ? (
            <div className="absolute inset-0 flex flex-col items-center justify-center text-center p-6 text-slate-400">
              <div className="w-20 h-20 bg-slate-800/80 rounded-full border-2 border-dashed border-slate-600 flex items-center justify-center text-3xl mb-4 shadow-xl">
                🏛️
              </div>
              <h3 className="text-xl font-bold text-slate-200 mb-1">
                Sala Weselna jest Pusta
              </h3>
              <p className="text-xs text-slate-400 max-w-md mb-6">
                Nie masz jeszcze żadnych stołów na sali. Dodaj nowy stół z
                lewego panelu lub zaimportuj listę gości z pliku Excel.
              </p>
              <div className="flex flex-wrap gap-3 justify-center">
                <label className="bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold px-4 py-2 rounded-lg cursor-pointer transition shadow">
                  📁 Importuj plik Excel (.xlsx)
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={handleExcelImport}
                    className="hidden"
                  />
                </label>
              </div>
            </div>
          ) : (
            <div
              style={{
                transform: `scale(${zoom})`,
                transformOrigin: "0 0",
                width: "2600px",
                height: "2600px",
              }}
              className="relative"
            >
              {tables.map((table) => {
                const tableGuests = guests.filter(
                  (g) => g.tableId === table.id,
                );
                const isSelected = selectedTableId === table.id;
                const isDragOverTable = dragOverTarget === table.id;
                const showFullNames =
                  isSelected || zoom >= 1.2 || alwaysShowNames;

                return (
                  <div
                    key={table.id}
                    onMouseDown={(e) => handleMouseDownTable(e, table)}
                    onClick={(e) => {
                      e.stopPropagation();
                      setSelectedTableId(table.id);
                    }}
                    onDragOver={(e) => handleDragOverTarget(e, table.id)}
                    onDragLeave={handleDragLeaveTarget}
                    onDrop={(e) => handleDropOnTable(e, table.id)}
                    style={{
                      left: `${table.x}px`,
                      top: `${table.y}px`,
                      transform: `rotate(${table.rotation || 0}deg)`,
                    }}
                    className={`absolute cursor-pointer transition-all duration-150 group ${
                      isDragOverTable
                        ? "ring-4 ring-emerald-400 ring-offset-4 ring-offset-slate-900 z-40 scale-105"
                        : isSelected
                          ? "ring-4 ring-amber-400 ring-offset-4 ring-offset-slate-900 z-30 scale-105"
                          : "z-10 hover:ring-2 hover:ring-amber-500/60"
                    }`}
                  >
                    <div
                      className={`relative flex flex-col items-center justify-center border-2 ${
                        isDragOverTable
                          ? "border-emerald-400 bg-emerald-950 text-emerald-100 shadow-2xl"
                          : isSelected
                            ? "border-amber-400 bg-linear-to-br from-amber-800 to-amber-950 text-amber-100 shadow-2xl"
                            : "border-amber-600/80 bg-linear-to-br from-amber-900/90 to-amber-950 text-amber-200 shadow-xl"
                      } ${table.shape === "round" ? "rounded-full w-40 h-40" : "rounded-2xl w-64 h-32"}`}
                    >
                      <div className="text-center px-2 pointer-events-none">
                        <div className="font-black text-sm text-amber-300 leading-tight uppercase tracking-wide">
                          {table.name}
                        </div>
                        <div className="text-[10px] text-amber-400/90 mt-0.5 font-medium">
                          miejsca: {table.capacity} ({tableGuests.length}{" "}
                          zajętych)
                        </div>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedTableId(table.id);
                            setIsDetailModalOpen(true);
                          }}
                          className="mt-1 bg-sky-500 hover:bg-sky-400 text-slate-950 text-[9px] font-black uppercase px-2 py-0.5 rounded-full shadow inline-flex items-center gap-1 opacity-90 hover:opacity-100 pointer-events-auto transition"
                        >
                          🔍 Widok Usadzenia
                        </button>
                      </div>

                      {/* Krzesełka wokół stołu */}
                      {Array.from({ length: table.capacity }).map((_, idx) => {
                        const guest = tableGuests.find(
                          (g) => g.seatIndex === idx,
                        );
                        const seatTargetId = `${table.id}-seat-${idx}`;
                        const isSeatDragOver = dragOverTarget === seatTargetId;

                        let xPercent = 0;
                        let yPercent = 0;

                        if (table.shape === "round") {
                          const angle = (idx / table.capacity) * (2 * Math.PI);
                          xPercent =
                            50 + (showFullNames ? 68 : 58) * Math.cos(angle);
                          yPercent =
                            50 + (showFullNames ? 68 : 58) * Math.sin(angle);
                        } else if (table.shape === "rectangle-one-side") {
                          const step =
                            table.capacity > 1 ? 85 / (table.capacity - 1) : 0;
                          xPercent = table.capacity === 1 ? 50 : 8 + idx * step;
                          yPercent = showFullNames ? -28 : -18;
                        } else {
                          const side = idx % 2 === 0 ? -1 : 1;
                          const posInSide = Math.floor(idx / 2);
                          const totalOnSide = Math.ceil(table.capacity / 2);
                          const step =
                            totalOnSide > 1 ? 85 / (totalOnSide - 1) : 0;
                          xPercent =
                            totalOnSide === 1 ? 50 : 8 + posInSide * step;
                          yPercent =
                            side === -1
                              ? showFullNames
                                ? -28
                                : -18
                              : showFullNames
                                ? 128
                                : 118;
                        }

                        return (
                          <div
                            key={idx}
                            draggable={!!guest}
                            onDragStart={(e) =>
                              guest && handleDragStartGuest(e, guest.id)
                            }
                            onDragOver={(e) =>
                              handleDragOverTarget(e, seatTargetId)
                            }
                            onDragLeave={handleDragLeaveTarget}
                            onDrop={(e) => handleDropOnSeat(e, table.id, idx)}
                            style={{
                              left: `${xPercent}%`,
                              top: `${yPercent}%`,
                              transform: "translate(-50%, -50%)",
                            }}
                            title={
                              guest
                                ? `Przeciągnij ${guest.fullName}`
                                : `Krzesło ${idx + 1} (wolne)`
                            }
                            className={`absolute flex items-center justify-center border transition-all shadow cursor-grab active:cursor-grabbing ${
                              isSeatDragOver
                                ? "bg-emerald-400 border-white ring-4 ring-emerald-300 z-50 text-slate-950 scale-125"
                                : guest
                                  ? showFullNames
                                    ? "bg-sky-400 border-sky-100 text-slate-900 px-2 py-1 rounded-xl text-[10px] font-bold shadow-xl z-40 whitespace-nowrap min-w-13.75 text-center border-2"
                                    : "w-7 h-7 rounded-full text-[10px] font-bold bg-sky-400 border-sky-200 text-slate-900"
                                  : "w-7 h-7 rounded-full text-[10px] font-bold bg-slate-800/90 border-slate-600 text-slate-500 hover:bg-slate-700"
                            }`}
                          >
                            {guest ? (
                              showFullNames ? (
                                <span className="truncate max-w-27.5 pointer-events-none">
                                  {guest.fullName}
                                </span>
                              ) : (
                                getInitials(guest.fullName)
                              )
                            ) : (
                              idx + 1
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {/* --- PRAWY PANEL: EDYCJA AKTYWNEGO STOŁU --- */}
      <div className="w-80 lg:w-88 bg-slate-800 border-l border-slate-700 flex flex-col h-full z-10 shrink-0">
        {activeTable ? (
          <div className="flex flex-col h-full">
            <div className="p-4 border-b border-slate-700 bg-slate-850">
              <div className="flex items-center justify-between mb-2">
                <h3 className="text-sm font-bold text-amber-400 truncate pr-2">
                  {activeTable.name}
                </h3>
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    onClick={() => setIsDetailModalOpen(true)}
                    className="p-1 bg-sky-600 hover:bg-sky-500 text-white rounded transition"
                    title="Otwórz widok zbliżenia"
                  >
                    <IconMaximize />
                  </button>
                  <button
                    onClick={() => handleRemoveTable(activeTable.id)}
                    className="p-1 bg-red-900/80 hover:bg-red-800 text-red-200 rounded transition"
                    title="Usuń ten stół"
                  >
                    <IconTrash />
                  </button>
                </div>
              </div>

              {/* USTAWIENIA PARAMETRÓW STOŁU */}
              <div className="space-y-2 mt-3 text-xs">
                <div>
                  <label className="text-[11px] text-slate-400 block mb-1">
                    Nazwa stołu:
                  </label>
                  <input
                    type="text"
                    value={activeTable.name}
                    onChange={(e) =>
                      handleUpdateTableProps(
                        activeTable.id,
                        "name",
                        e.target.value,
                      )
                    }
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-white focus:outline-none focus:border-indigo-500"
                  />
                </div>

                <div className="flex gap-2">
                  <div className="flex-1">
                    <label className="text-[11px] text-slate-400 block mb-1">
                      Kształt:
                    </label>
                    <select
                      value={activeTable.shape}
                      onChange={(e) =>
                        handleUpdateTableProps(
                          activeTable.id,
                          "shape",
                          e.target.value as any,
                        )
                      }
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-white focus:outline-none"
                    >
                      <option value="round">Okrągły 🟣</option>
                      <option value="rectangle">
                        Prostokątny (2 strony) 🟦
                      </option>
                      <option value="rectangle-one-side">
                        Prostokątny (1 strona) ▭
                      </option>
                    </select>
                  </div>

                  <div className="w-20">
                    <label className="text-[11px] text-slate-400 block mb-1">
                      Miejsca:
                    </label>
                    <input
                      type="number"
                      min="1"
                      max="50"
                      value={activeTable.capacity}
                      onChange={(e) =>
                        handleUpdateTableProps(
                          activeTable.id,
                          "capacity",
                          parseInt(e.target.value) || 1,
                        )
                      }
                      className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-white text-center focus:outline-none"
                    />
                  </div>
                </div>

                <div>
                  <label className="text-[11px] text-slate-400 block mb-1">
                    Obrót ({activeTable.rotation || 0}°):
                  </label>
                  <input
                    type="range"
                    min="0"
                    max="360"
                    step="15"
                    value={activeTable.rotation || 0}
                    onChange={(e) =>
                      handleUpdateTableProps(
                        activeTable.id,
                        "rotation",
                        parseInt(e.target.value) || 0,
                      )
                    }
                    className="w-full accent-indigo-500 cursor-pointer"
                  />
                </div>
              </div>
            </div>

            {/* SZYBKIE DODAWANIE GOŚCIA DO TEGO STOŁU */}
            <div className="p-3 border-b border-slate-700 bg-slate-800">
              <label className="text-[11px] font-medium text-slate-300 block mb-1">
                Szybkie dodanie osoby do tego stołu:
              </label>
              <div className="flex gap-1.5">
                <input
                  type="text"
                  placeholder="Imię i nazwisko"
                  value={quickGuestName}
                  onChange={(e) => setQuickGuestName(e.target.value)}
                  onKeyDown={(e) =>
                    e.key === "Enter" && handleAddQuickGuestToSelectedTable()
                  }
                  className="flex-1 bg-slate-900 border border-slate-700 rounded px-2.5 py-1 text-xs text-white focus:outline-none focus:border-indigo-500"
                />
                <button
                  onClick={handleAddQuickGuestToSelectedTable}
                  className="bg-indigo-600 hover:bg-indigo-500 text-white px-2.5 py-1 rounded text-xs font-bold transition"
                >
                  Dodaj
                </button>
              </div>
            </div>

            {/* LISTA MIEJSC W TYM STOLE */}
            <div className="flex-1 overflow-y-auto p-4 space-y-2">
              <h4 className="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2 flex justify-between">
                <span>Lista miejsc ({activeTable.capacity})</span>
              </h4>

              {Array.from({ length: activeTable.capacity }).map((_, idx) => {
                const assignedGuest = guests.find(
                  (g) => g.tableId === activeTable.id && g.seatIndex === idx,
                );

                return (
                  <div
                    key={idx}
                    onDragOver={(e) =>
                      handleDragOverTarget(
                        e,
                        `${activeTable.id}-panel-seat-${idx}`,
                      )
                    }
                    onDragLeave={handleDragLeaveTarget}
                    onDrop={(e) => handleDropOnSeat(e, activeTable.id, idx)}
                    className={`border rounded-md p-2 flex justify-between items-center text-xs transition-colors ${
                      dragOverTarget === `${activeTable.id}-panel-seat-${idx}`
                        ? "bg-emerald-950 border-emerald-400"
                        : "bg-slate-700/40 border-slate-600/60"
                    }`}
                  >
                    <div className="flex items-center gap-2 truncate pr-2">
                      <span className="w-5 h-5 rounded-full bg-slate-800 text-slate-400 text-[10px] font-bold flex items-center justify-center border border-slate-600 shrink-0">
                        {idx + 1}
                      </span>
                      {assignedGuest ? (
                        <span
                          draggable
                          onDragStart={(e) =>
                            handleDragStartGuest(e, assignedGuest.id)
                          }
                          className="font-bold text-sky-300 truncate cursor-grab active:cursor-grabbing hover:underline"
                          title="Przeciągnij w inne miejsce"
                        >
                          ✋ {assignedGuest.fullName}
                        </span>
                      ) : (
                        <span className="text-slate-500 italic text-[11px]">
                          Wolne krzesło
                        </span>
                      )}
                    </div>

                    {assignedGuest ? (
                      <button
                        onClick={() =>
                          handleAssignGuestToTableSeat(
                            assignedGuest.id,
                            null,
                            null,
                          )
                        }
                        className="text-red-400 hover:text-red-300 text-[11px] hover:underline shrink-0"
                      >
                        Odstaw
                      </button>
                    ) : (
                      <span className="text-[10px] text-slate-500 font-mono">
                        Puste
                      </span>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        ) : (
          <div className="flex-1 flex flex-col items-center justify-center text-center p-6 text-slate-500 text-xs gap-3">
            <div className="text-3xl">🪑</div>
            <p>
              Kliknij na dowolny stół na sali, aby zobaczyć szczegóły i opcje
              usadzania gości.
            </p>
          </div>
        )}
      </div>

      {/* --- MODAL: WIDOK ZBLIŻENIA STOŁU --- */}
      {isDetailModalOpen && activeTable && (
        <div className="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md flex flex-col justify-between p-4 md:p-8 animate-fadeIn overflow-hidden">
          <div className="flex items-center justify-between border-b border-slate-700/80 pb-4 mb-4">
            <div className="flex items-center gap-3">
              <span className="text-2xl">🪑</span>
              <div>
                <h2 className="text-xl font-black text-white flex items-center gap-2">
                  Widok Usadzenia:{" "}
                  <span className="text-sky-400 underline">
                    {activeTable.name}
                  </span>
                </h2>
                <p className="text-xs text-slate-400">
                  Przeciągaj kafelki gości bezpośrednio na wybrane krzesła lub
                  kliknij, aby zamienić miejsca.
                </p>
              </div>
            </div>

            <div className="flex items-center gap-3">
              {selectedSeatForSwap !== null && (
                <span className="bg-amber-500 text-slate-950 px-3 py-1 rounded-full text-xs font-bold animate-pulse">
                  Wybierz drugie krzesło, aby zamienić miejsca
                </span>
              )}

              <button
                onClick={() => {
                  setIsDetailModalOpen(false);
                  setSelectedSeatForSwap(null);
                }}
                className="bg-slate-800 hover:bg-slate-700 text-slate-200 p-2 rounded-xl border border-slate-600 transition flex items-center gap-1 text-xs font-bold"
              >
                <IconClose />
                <span>Zamknij Zbliżenie</span>
              </button>
            </div>
          </div>

          <div className="flex-1 bg-slate-300 rounded-2xl p-6 shadow-2xl flex flex-col justify-center items-center overflow-x-auto border-4 border-slate-400 relative">
            {activeTable.shape === "rectangle-one-side" ? (
              <div className="flex flex-col items-stretch w-full max-w-7xl space-y-3 min-w-225 select-none">
                <div className="grid grid-flow-col auto-cols-fr gap-2 px-2">
                  {Array.from({ length: activeTable.capacity }).map(
                    (_, colIdx) => {
                      const seatIndex = colIdx;
                      const guest = activeTableGuestsMap[seatIndex];
                      const seatKey = `${activeTable.id}-modalseat-${seatIndex}`;
                      const isTargetOver = dragOverTarget === seatKey;

                      return (
                        <div
                          
                          key={`one-side-seat-${seatIndex}`}
                          
                          title={ guest.fullName}
                          draggable={!!guest}
                          onDragStart={(e) =>
                            guest && handleDragStartGuest(e, guest.id)
                          }
                          onDragOver={(e) => handleDragOverTarget(e, seatKey)}
                          onDragLeave={handleDragLeaveTarget}
                          onDrop={(e) =>
                            handleDropOnSeat(e, activeTable.id, seatIndex)
                          }
                          onClick={() => {
                        
                            if (selectedSeatForSwap === null) {
                              if (guest) setSelectedSeatForSwap(seatIndex);
                            } else {
                              const guest1 =
                                activeTableGuestsMap[selectedSeatForSwap];
                              if (guest1) {
                                handleAssignGuestToTableSeat(
                                  guest1.id,
                                  activeTable.id,
                                  seatIndex,
                                );
                              }
                              setSelectedSeatForSwap(null);
                            }
                          }}
                          className={`h-16 rounded-2xl flex items-center justify-center p-2 text-center text-xs md:text-sm font-semibold transition-all duration-150 shadow-md cursor-grab active:cursor-grabbing border-2 ${
                            isTargetOver
                              ? "bg-emerald-400 text-slate-950 border-white ring-4 ring-emerald-300 scale-105 z-30 font-black"
                              : selectedSeatForSwap === seatIndex
                                ? "bg-amber-400 text-slate-950 border-amber-600 ring-4 ring-amber-300 scale-105 z-20 font-black"
                                : guest
                                  ? "bg-sky-400 hover:bg-sky-300 text-slate-900 border-sky-300 hover:scale-102 shadow-sky-400/20"
                                  : "bg-slate-100/90 hover:bg-white text-slate-400 border-slate-300 border-dashed"
                          }`}
                        >
                          <span className="truncate max-w-full pointer-events-none">
                            {guest
                              ? guest.fullName
                              : `krzesło ${seatIndex + 1}`}
                          </span>
                        </div>
                      );
                    },
                  )}
                </div>

                <div className="h-24 bg-slate-200 border-2 border-slate-400 flex items-center justify-between px-6 shadow-inner rounded-xl">
                  <div className="text-slate-500 font-mono text-xs">
                    miejsca: {activeTable.capacity} (jednostronne)
                  </div>

                  <div className="text-slate-700 font-black text-xl md:text-2xl uppercase tracking-widest text-center flex-1">
                    {activeTable.name}
                  </div>

                  <div className="text-slate-500 font-mono text-xs">
                    zajęte: {Object.keys(activeTableGuestsMap).length}
                  </div>
                </div>
              </div>
            ) : activeTable.shape === "rectangle" ? (
              <div className="flex flex-col items-stretch w-full max-w-7xl space-y-3 min-w-225 select-none">
                <div className="grid grid-flow-col auto-cols-fr gap-2 px-2">
                  {Array.from({
                    length: Math.ceil(activeTable.capacity / 2),
                  }).map((_, colIdx) => {
                    const seatIndex = colIdx;
                    const guest = activeTableGuestsMap[seatIndex];
                    const seatKey = `${activeTable.id}-modalseat-${seatIndex}`;
                    const isTargetOver = dragOverTarget === seatKey;

                    return (
                      <div
                        key={`top-seat-${seatIndex}`}
                        title={guest?.fullName||guest}
                        draggable={!!guest}
                        onDragStart={(e) =>
                          guest && handleDragStartGuest(e, guest.id)
                        }
                        onDragOver={(e) => handleDragOverTarget(e, seatKey)}
                        onDragLeave={handleDragLeaveTarget}
                        onDrop={(e) =>
                          handleDropOnSeat(e, activeTable.id, seatIndex)
                        }
                        onClick={() => {
                          if (selectedSeatForSwap === null) {
                            if (guest) setSelectedSeatForSwap(seatIndex);
                          } else {
                            const guest1 =
                              activeTableGuestsMap[selectedSeatForSwap];
                            if (guest1) {
                              handleAssignGuestToTableSeat(
                                guest1.id,
                                activeTable.id,
                                seatIndex,
                              );
                            }
                            setSelectedSeatForSwap(null);
                          }
                        }}
                        className={`h-16 rounded-2xl flex items-center justify-center p-2 text-center text-xs md:text-sm font-semibold transition-all duration-150 shadow-md cursor-grab active:cursor-grabbing border-2 ${
                          isTargetOver
                            ? "bg-emerald-400 text-slate-950 border-white ring-4 ring-emerald-300 scale-105 z-30 font-black"
                            : selectedSeatForSwap === seatIndex
                              ? "bg-amber-400 text-slate-950 border-amber-600 ring-4 ring-amber-300 scale-105 z-20 font-black"
                              : guest
                                ? "bg-sky-400 hover:bg-sky-300 text-slate-900 border-sky-300 hover:scale-102 shadow-sky-400/20"
                                : "bg-slate-100/90 hover:bg-white text-slate-400 border-slate-300 border-dashed"
                        }`}
                      >
                        <span className="truncate max-w-full pointer-events-none">
                          {guest ? guest.fullName : `krzesło ${seatIndex + 1}`}
                        </span>
                      </div>
                    );
                  })}
                </div>

                <div className="h-20 bg-slate-200 border-y-2 border-slate-400 flex items-center justify-between px-6 shadow-inner rounded-md">
                  <div className="text-slate-500 font-mono text-xs">
                    miejsca: {activeTable.capacity}
                  </div>

                  <div className="text-slate-700 font-black text-xl md:text-2xl uppercase tracking-widest text-center flex-1">
                    {activeTable.name}
                  </div>

                  <div className="text-slate-500 font-mono text-xs">
                    zajęte: {Object.keys(activeTableGuestsMap).length}
                  </div>
                </div>

                <div className="grid grid-flow-col auto-cols-fr gap-2 px-2">
                  {Array.from({
                    length: Math.floor(activeTable.capacity / 2),
                  }).map((_, colIdx) => {
                    const topHalfCount = Math.ceil(activeTable.capacity / 2);
                    const seatIndex = topHalfCount + colIdx;
                    const guest = activeTableGuestsMap[seatIndex];
                    const seatKey = `${activeTable.id}-modalseat-${seatIndex}`;
                    const isTargetOver = dragOverTarget === seatKey;

                    return (
                      <div
                        key={`bottom-seat-${seatIndex}`}
                        
                        title={guest?.fullName||guest}
                        draggable={!!guest}
                        onDragStart={(e) =>
                          guest && handleDragStartGuest(e, guest.id)
                        }
                        onDragOver={(e) => handleDragOverTarget(e, seatKey)}
                        onDragLeave={handleDragLeaveTarget}
                        onDrop={(e) =>
                          handleDropOnSeat(e, activeTable.id, seatIndex)
                        }
                        onClick={() => {
                          if (selectedSeatForSwap === null) {
                            if (guest) setSelectedSeatForSwap(seatIndex);
                          } else {
                            const guest1 =
                              activeTableGuestsMap[selectedSeatForSwap];
                            if (guest1) {
                              handleAssignGuestToTableSeat(
                                guest1.id,
                                activeTable.id,
                                seatIndex,
                              );
                            }
                            setSelectedSeatForSwap(null);
                          }
                        }}
                        className={`h-16 rounded-2xl flex items-center justify-center p-2 text-center text-xs md:text-sm font-semibold transition-all duration-150 shadow-md cursor-grab active:cursor-grabbing border-2 ${
                          isTargetOver
                            ? "bg-emerald-400 text-slate-950 border-white ring-4 ring-emerald-300 scale-105 z-30 font-black"
                            : selectedSeatForSwap === seatIndex
                              ? "bg-amber-400 text-slate-950 border-amber-600 ring-4 ring-amber-300 scale-105 z-20 font-black"
                              : guest
                                ? "bg-sky-400 hover:bg-sky-300 text-slate-900 border-sky-300 hover:scale-102 shadow-sky-400/20"
                                : "bg-slate-100/90 hover:bg-white text-slate-400 border-slate-300 border-dashed"
                        }`}
                      >
                        <span className="truncate max-w-full pointer-events-none">
                          {guest ? guest.fullName : `krzesło ${seatIndex + 1}`}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (
              <div className="relative w-[500px] h-125 flex items-center justify-center select-none">
                <div className="w-64 h-64 bg-slate-200 border-4 border-slate-400 rounded-full flex flex-col items-center justify-center text-center p-4 shadow-inner z-10">
                  <span className="text-slate-800 font-black text-xl uppercase tracking-wider">
                    {activeTable.name}
                  </span>
                  <span className="text-slate-500 font-mono text-xs mt-1">
                    miejsca: {activeTable.capacity}
                  </span>
                </div>

                {Array.from({ length: activeTable.capacity }).map(
                  (_, seatIndex) => {
                    const angle =
                      (seatIndex / activeTable.capacity) * (2 * Math.PI) -
                      Math.PI / 2;
                    const radius = 200;
                    const x = Math.cos(angle) * radius;
                    const y = Math.sin(angle) * radius;

                    const guest = activeTableGuestsMap[seatIndex];
                    const seatKey = `${activeTable.id}-modalseat-${seatIndex}`;
                    const isTargetOver = dragOverTarget === seatKey;
                    

                    return (
                      <div
                        key={`round-seat-${seatIndex}`}
                        style={{
                          transform: `translate(${x}px, ${y}px)`,
                        }}
                     
                        draggable={!!guest}
                        onDragStart={(e) =>
                          guest && handleDragStartGuest(e, guest.id)
                        }
                        onDragOver={(e) => handleDragOverTarget(e, seatKey)}
                        onDragLeave={handleDragLeaveTarget}
                        onDrop={(e) =>
                          handleDropOnSeat(e, activeTable.id, seatIndex)
                        }
                        onClick={() => {
                          if (selectedSeatForSwap === null) {
                            if (guest) setSelectedSeatForSwap(seatIndex);
                          } else {
                            const guest1 =
                              activeTableGuestsMap[selectedSeatForSwap];
                            if (guest1) {
                              handleAssignGuestToTableSeat(
                                guest1.id,
                                activeTable.id,
                                seatIndex,
                              );
                            }
                            setSelectedSeatForSwap(null);
                          }
                        }}
                        className={`absolute w-28 h-14 rounded-2xl flex items-center justify-center p-2 text-center text-xs font-bold transition-all duration-150 shadow-lg cursor-grab active:cursor-grabbing border-2 z-20 ${
                          isTargetOver
                            ? "bg-emerald-400 text-slate-950 border-white ring-4 ring-emerald-300 scale-110 font-black z-30"
                            : selectedSeatForSwap === seatIndex
                              ? "bg-amber-400 text-slate-950 border-amber-600 ring-4 ring-amber-300 scale-110 font-black"
                              : guest
                                ? "bg-sky-400 hover:bg-sky-300 text-slate-900 border-sky-300"
                                : "bg-slate-100/90 hover:bg-white text-slate-400 border-slate-300 border-dashed"
                        }`}
                      >
                        <span className="truncate max-w-full pointer-events-none">
                          {guest ? guest.fullName : `krzesło ${seatIndex + 1}`}
                        </span>
                      </div>
                    );
                  },
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}