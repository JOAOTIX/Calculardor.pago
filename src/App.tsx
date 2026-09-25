import { useState, useMemo } from 'react'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import toast, { Toaster } from 'react-hot-toast'
import './App.css'

// Tipos de datos
interface DiaData {
  id: number;
  dia: string;
}

interface RegistroDiario {
  id: string;
  dia: string; 
  horaEntrada: string;
  horaSalida: string;
}

interface SemanaData {
  id: string; 
  titulo: string;
  registros: RegistroDiario[];
}

function App() {
  // Configuración general
  const [pagoPorHora, setPagoPorHora] = useState<number | ''>('');

  // Arreglo de días dinámicos
  const [dias] = useState<DiaData[]>([
    { id: 1, dia: 'Lunes' },
    { id: 2, dia: 'Martes' },
    { id: 3, dia: 'Miércoles' },
    { id: 4, dia: 'Jueves' },
    { id: 5, dia: 'Viernes' },
    { id: 6, dia: 'Sábado' },
    { id: 7, dia: 'Domingo' },
  ]);
  
  // Estado para las 3 semanas iniciales
  const [semanas, setSemanas] = useState<SemanaData[]>([
    { id: crypto.randomUUID(), titulo: "SEMANA DEL 10-15 DE AGOSTO", registros: [{ id: crypto.randomUUID(), dia: 'Lunes', horaEntrada: '', horaSalida: '' }] },
  ]);

  // Función para calcular horas en formato decimal
  const calcularHorasDecimal = (entrada: string, salida: string): number => {
    if (!entrada || !salida) return 0;
    const [entHora, entMin] = entrada.split(':').map(Number);
    const [salHora, salMin] = salida.split(':').map(Number);
    
    const fechaEntrada = new Date(0, 0, 0, entHora, entMin, 0);
    const fechaSalida = new Date(0, 0, 0, salHora, salMin, 0);
    
    if (fechaSalida < fechaEntrada) fechaSalida.setDate(fechaSalida.getDate() + 1);
    
    return (fechaSalida.getTime() - fechaEntrada.getTime()) / (1000 * 60 * 60);
  };

  // Convertir decimal a formato HH:MM
  const formatoHorasMinutos = (decimalHoras: number, paraExcel = false) => {
    if (decimalHoras === 0) return paraExcel ? "00:00:00" : "--:--";
    const h = Math.floor(decimalHoras);
    const m = Math.round((decimalHoras - h) * 60);
    
    const horasStr = h.toString().padStart(2, '0');
    const minStr = m.toString().padStart(2, '0');
    
    if (paraExcel) {
      return `${horasStr}:${minStr}:00`;
    }
    return `${horasStr}h ${minStr}m`;
  };

  // -- FUNCIÓN: Exportar a Excel --
  const exportarAExcel = async () => {
    let totalGeneralHorasDecimal = 0;
    
    semanas.forEach(semana => {
      semana.registros.forEach(reg => {
        if (reg.dia && reg.horaEntrada && reg.horaSalida) {
          totalGeneralHorasDecimal += calcularHorasDecimal(reg.horaEntrada, reg.horaSalida);
        }
      });
    });

    if (totalGeneralHorasDecimal === 0) {
      toast.error('No hay dias seleccionados para exportar. Por favor, llena los campos.', {
        style: { borderRadius: '10px', background: '#333', color: '#fff', padding: '16px', fontWeight: 'bold' },
        duration: 4000,
      });
      return;
    }

    if (!pagoPorHora) {
      toast.error('Recuerda ingresar la tarifa por hora antes de exportar.', {
        style: { borderRadius: '10px', background: '#333', color: '#fff', padding: '16px', fontWeight: 'bold' },
        duration: 4000,
      });
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Anahí', { views: [{ showGridLines: false }] });

    sheet.columns = [
      { width: 22 }, // Día
      { width: 15 }, // Ingreso
      { width: 15 }, // Salida
      { width: 20 }  // Horas trabajadas
    ];

    const borderStyle: Partial<ExcelJS.Borders> = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
    const centerAlign: Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'center', wrapText: true };

    semanas.forEach(semana => {
      // 1. Título de la semana (Morado)
      const titleRow = sheet.addRow([semana.titulo]);
      sheet.mergeCells(`A${titleRow.number}:D${titleRow.number}`);
      titleRow.height = 30; // Dar un poco más de alto por si hay 2 líneas
      for(let i = 1; i <= 4; i++) {
        const cell = titleRow.getCell(i);
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C3AED' } }; 
        cell.border = borderStyle;
        cell.alignment = centerAlign;
      }

      // 2. Subcabeceras (Lila pastel)
      const headerRow = sheet.addRow(['Día', 'Ingreso', 'Salida', 'Horas trabajadas']);
      for(let i = 1; i <= 4; i++) {
        const cell = headerRow.getCell(i);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9D5FF' } }; 
        cell.font = { color: { argb: 'FF3B0764' }, bold: true }; 
        cell.border = borderStyle;
        cell.alignment = centerAlign;
      }

      let totalSemanaDecimal = 0;
      let tieneRegistros = false;

      // 3. Registros diarios
      semana.registros.forEach(reg => {
        if (reg.dia && reg.horaEntrada && reg.horaSalida) {
          tieneRegistros = true;
          const horasDecimales = calcularHorasDecimal(reg.horaEntrada, reg.horaSalida);
          totalSemanaDecimal += horasDecimales;
          
          const dataRow = sheet.addRow([
            reg.dia, 
            reg.horaEntrada, 
            reg.horaSalida, 
            formatoHorasMinutos(horasDecimales, true)
          ]);
          
          for(let i = 1; i <= 4; i++) {
            const cell = dataRow.getCell(i);
            cell.border = borderStyle;
            cell.alignment = centerAlign;
          }
        }
      });
      
      // Fila vacía si no hay registros
      if (!tieneRegistros) {
        const emptyRow = sheet.addRow(['', '', '', '']);
        for(let i = 1; i <= 4; i++) {
          const cell = emptyRow.getCell(i);
          cell.border = borderStyle;
          cell.alignment = centerAlign;
        }
      }

      // 4. Fila de Total de la Semana
      const totalRow = sheet.addRow(['', '', 'TOTAL SEMANA:', formatoHorasMinutos(totalSemanaDecimal, true)]);
      for(let i = 1; i <= 4; i++) {
        const cell = totalRow.getCell(i);
        cell.border = borderStyle;
        cell.alignment = centerAlign;
      }
      totalRow.getCell(3).font = { bold: true };
      totalRow.getCell(4).font = { bold: true, color: { argb: 'FF00B050' } }; // Verde

      // Espaciado entre semanas
      sheet.addRow([]);
      sheet.addRow([]);
    });

    // --- SECCIÓN RESUMEN FINAL ---
    const resumenTitle = sheet.addRow(['RESUMEN FINAL']);
    sheet.mergeCells(`A${resumenTitle.number}:D${resumenTitle.number}`);
    for(let i = 1; i <= 4; i++) {
      const cell = resumenTitle.getCell(i);
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D935C' } }; 
      cell.border = borderStyle;
      cell.alignment = centerAlign;
    }

    const rowHoras = sheet.addRow(['Total Horas', '', '', formatoHorasMinutos(totalGeneralHorasDecimal, true)]);
    const rowTarifa = sheet.addRow(['Tarifa por Hora', '', '', `S/. ${pagoPorHora}`]);
    const pagoTotal = totalGeneralHorasDecimal * (Number(pagoPorHora) || 0);
    const rowCobrar = sheet.addRow(['TOTAL A COBRAR', '', '', `S/. ${pagoTotal.toFixed(2)}`]);

    [rowHoras, rowTarifa, rowCobrar].forEach(row => {
      for(let i = 1; i <= 4; i++) {
        const cell = row.getCell(i);
        cell.border = borderStyle;
        cell.alignment = centerAlign;
      }
    });

    rowHoras.getCell(1).font = { bold: true };
    rowHoras.getCell(4).font = { bold: true };
    
    rowCobrar.getCell(1).font = { bold: true, size: 11 };
    rowCobrar.getCell(4).font = { bold: true, size: 11, color: { argb: 'FF00B050' } };

    // Generar archivo
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'CALCULADOR_DE_HORAS.xlsx');

    toast.success('¡Archivo Excel generado con éxito!', {
      style: { borderRadius: '10px', background: '#059669', color: '#fff', padding: '16px', fontWeight: 'bold' },
    });
  };

  // --- ACCIONES: Manejo de Semanas ---
  const agregarSemana = () => {
    if (semanas.length >= 4) {
      toast.error('Solo se pueden agregar hasta 4 semanas.', { duration: 3000 });
      return;
    }
    
    const nuevaSemana: SemanaData = {
      id: crypto.randomUUID(),
      titulo: `SEMANA ${semanas.length + 1}`,
      registros: [{ id: crypto.randomUUID(), dia: '', horaEntrada: '', horaSalida: '' }]
    };
    
    setSemanas([...semanas, nuevaSemana]);
  };

  const eliminarSemana = (idSemana: string) => {
    if (semanas.length === 1) {
      toast.error('Debe haber al menos una semana.', { duration: 2000 });
      return;
    }
    setSemanas(semanas.filter(s => s.id !== idSemana));
  };

  const actualizarTituloSemana = (semanaId: string, nuevoTitulo: string) => {
    setSemanas(semanas.map(s => s.id === semanaId ? { ...s, titulo: nuevoTitulo } : s));
  };

  // Ajusta la altura del Textarea cuando React lo dibuja en pantalla o cuando escribes
  const ajustarAltura = (el: HTMLTextAreaElement | null) => {
    if (el) {
      el.style.height = 'auto'; // Resetea la altura
      el.style.height = `${el.scrollHeight}px`; // Asigna la altura exacta del texto
    }
  };

  // Acciones de Días dentro de una Semana
  const agregarDia = (semanaId: string) => {
    setSemanas(semanas.map(s => {
      if (s.id === semanaId) {
        // Validar que no haya más de 7 días
        if (s.registros.length >= 7) {
          toast.error('No puedes agregar más de 7 días a una semana.', { duration: 3000 });
          return s; // Devolver la semana sin cambios
        }
        return { 
          ...s, 
          registros: [...s.registros, { id: crypto.randomUUID(), dia: '', horaEntrada: '', horaSalida: '' }] 
        };
      }
      return s;
    }));
  };

  const actualizarRegistro = (semanaId: string, idRegistro: string, campo: keyof RegistroDiario, valor: string) => {
    setSemanas(semanas.map(s => 
      s.id === semanaId 
        ? { ...s, registros: s.registros.map(r => r.id === idRegistro ? { ...r, [campo]: valor } : r) }
        : s
    ));
  };

  const eliminarRegistro = (semanaId: string, idRegistro: string) => {
    setSemanas(semanas.map(s => 
      s.id === semanaId 
        ? { ...s, registros: s.registros.filter(r => r.id !== idRegistro) }
        : s
    ));
  };

  // Cálculos por semana y generales
  const subtotales = useMemo(() => {
    return semanas.map(semana => {
      let horasDecimal = 0;
      semana.registros.forEach(reg => {
        horasDecimal += calcularHorasDecimal(reg.horaEntrada, reg.horaSalida);
      });
      return { 
        id: semana.id, 
        horasDecimal, 
        pago: horasDecimal * (Number(pagoPorHora) || 0) 
      };
    });
  }, [semanas, pagoPorHora]);

  const totalGeneral = useMemo(() => {
    const totalHrs = subtotales.reduce((acc, curr) => acc + curr.horasDecimal, 0);
    const totalPgo = subtotales.reduce((acc, curr) => acc + curr.pago, 0);
    return { horasDecimal: totalHrs, pago: totalPgo };
  }, [subtotales]);

  return (
    <main className="min-h-screen bg-[#f3f4f6] pb-32 font-sans text-gray-800">
      <Toaster position="top-center" reverseOrder={false} />

      {/* Header Fijo Superior */}
      <div className="bg-white shadow-sm py-4 px-4 sticky top-0 z-40 mb-6 flex justify-between items-center">
         <div className="flex items-center gap-2 md:gap-3">
            <svg width="28" height="28" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg" className="drop-shadow-sm md:w-8 md:h-8">
              <rect x="5" y="5" width="90" height="90" rx="22" fill="#EEF2FF" stroke="#DFE7FF" strokeWidth="2"/>
              <rect x="37" y="14" width="16" height="6" rx="2" fill="#2563EB" />
              <circle cx="45" cy="47" r="22" stroke="#2563EB" strokeWidth="7" />
              <path d="M45 32 V47 L55 57" stroke="#2563EB" strokeWidth="5" strokeLinecap="round" strokeLinejoin="round" />
              <circle cx="45" cy="47" r="4" fill="#2563EB" />
              <circle cx="70" cy="70" r="17" fill="#059669" stroke="#EEF2FF" strokeWidth="4" />
              <text x="70" y="76" fontFamily="sans-serif" fontWeight="900" fontSize="16" fill="#FFFFFF" textAnchor="middle">S/</text>
            </svg>
            <h1 className="text-lg md:text-2xl font-black text-blue-600 tracking-tight">Calculadora</h1>
         </div>
         <button 
           onClick={exportarAExcel}
           className="bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-3 md:px-4 rounded-lg shadow flex items-center gap-1 md:gap-2 text-xs md:text-sm transition-colors"
         >
           <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 md:h-5 md:w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
             <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
           </svg>
           Exportar
         </button>
      </div>

      <div className="max-w-6xl mx-auto px-4">
        {/* Configuración Global (Pago) */}
        <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 max-w-sm mx-auto mb-8">
          <label className="block text-sm font-bold text-gray-700 mb-2 uppercase tracking-wide">
            Tarifa por hora (S/.)
          </label>
          <div className="relative">
            <span className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400 font-bold">S/ </span>
            <input 
              type="number" 
              value={pagoPorHora}
              onChange={(e) => setPagoPorHora(e.target.value ? Number(e.target.value) : '')}
              className="w-full rounded-lg bg-gray-50 border-gray-200 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 outline-none transition-all pl-8 p-3 text-lg font-semibold text-gray-800 border"
              placeholder="0.00"
            />
          </div>
        </div>

        {/* Renderizado de las Semanas */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {semanas.map((semana) => {
            const subtotal = subtotales.find(s => s.id === semana.id);
            
            return (
              <div key={semana.id} className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden flex flex-col h-full relative group/semana">
                
                {/* Cabecera de la Semana */}
                <div className="bg-gray-50 p-4 border-b border-gray-100 pt-5">
                  <div className="flex justify-between items-start gap-3 mb-3">
                    
                    {/* ✅ Se aplica ref={ajustarAltura} aquí */}
                    <textarea 
                      ref={ajustarAltura}
                      value={semana.titulo}
                      onChange={(e) => {
                        actualizarTituloSemana(semana.id, e.target.value);
                        ajustarAltura(e.target);
                      }}
                      rows={1}
                      className="flex-1 bg-transparent font-black text-gray-800 outline-none placeholder-gray-400 sm:text-lg break-words resize-none overflow-hidden"
                      placeholder="Nombre de la semana"
                      style={{ minHeight: '32px' }}
                    />
                    
                    <button 
                      onClick={() => eliminarSemana(semana.id)}
                      className="text-red-400 hover:text-red-600 bg-white hover:bg-red-50 rounded-md p-1.5 transition-colors border border-gray-200 shadow-sm flex-shrink-0"
                      title="Eliminar semana completa"
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                        <path fillRule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 012 0v6a1 1 0 11-2 0V8zm5-1a1 1 0 00-1 1v6a1 1 0 102 0V8a1 1 0 00-1-1z" clipRule="evenodd" />
                      </svg>
                    </button>

                  </div>
                  
                  <div className="flex justify-between items-center bg-white px-3 py-2 rounded-lg shadow-sm border border-gray-100 text-sm">
                    <span className="font-bold text-gray-500">SUBTOTAL</span> 
                    <span className="font-medium text-gray-700">
                      {formatoHorasMinutos(subtotal?.horasDecimal || 0)} 
                      <span className="text-gray-300 mx-2">|</span> 
                      <span className="text-blue-600 font-black">S/. {subtotal?.pago.toFixed(2)}</span>
                    </span>
                  </div>
                </div>
                
                {/* Lista de Registros */}
                <div className="p-4 space-y-4 flex-1">
                  {semana.registros.map((registro) => {
                    const horasDiarias = calcularHorasDecimal(registro.horaEntrada, registro.horaSalida);
                    
                    return (
                      <div key={registro.id} className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm hover:border-blue-300 transition-colors group relative">
                        <button 
                          onClick={() => eliminarRegistro(semana.id, registro.id)}
                          className="absolute -top-2 -right-2 bg-red-100 text-red-500 hover:bg-red-500 hover:text-white w-7 h-7 flex items-center justify-center rounded-full text-sm font-bold shadow-sm transition-colors opacity-80 group-hover:opacity-100"
                        >
                          ×
                        </button>

                        <div className="flex justify-between items-center mb-4">
                          <select 
                            value={registro.dia}
                            onChange={(e) => actualizarRegistro(semana.id, registro.id, 'dia', e.target.value)}
                            className="w-1/2 font-bold text-gray-700 outline-none border-b border-transparent focus:border-blue-500 bg-transparent cursor-pointer appearance-none"
                          >
                            <option value="" disabled>Seleccionar día...</option>
                            {dias.map((d) => (
                              <option key={d.id} value={d.dia}>
                                {d.dia}
                              </option>
                            ))}
                          </select>
                          
                          <span className="text-xs font-bold text-blue-700 bg-blue-50 px-2 py-1 rounded-md border border-blue-100">
                            {formatoHorasMinutos(horasDiarias)}
                          </span>
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="text-[10px] uppercase tracking-wider text-gray-400 font-bold mb-1.5 block">Ingreso</label>
                            <input 
                              type="time" 
                              value={registro.horaEntrada}
                              onChange={(e) => actualizarRegistro(semana.id, registro.id, 'horaEntrada', e.target.value)}
                              className="w-full bg-gray-50 border border-gray-200 rounded-lg p-2 text-sm text-gray-700 outline-none focus:border-blue-500 focus:bg-white transition-colors"
                            />
                          </div>
                          <div>
                            <label className="text-[10px] uppercase tracking-wider text-gray-400 font-bold mb-1.5 block">Salida</label>
                            <input 
                              type="time" 
                              value={registro.horaSalida}
                              onChange={(e) => actualizarRegistro(semana.id, registro.id, 'horaSalida', e.target.value)}
                              className="w-full bg-gray-50 border border-gray-200 rounded-lg p-2 text-sm text-gray-700 outline-none focus:border-blue-500 focus:bg-white transition-colors"
                            />
                          </div>
                        </div>
                      </div>
                    );
                  })}
                  
                  {semana.registros.length < 7 && (
                    <button 
                      onClick={() => agregarDia(semana.id)}
                      className="w-full py-3 mt-2 rounded-xl text-sm font-bold text-gray-500 bg-gray-50 border-2 border-dashed border-gray-200 hover:border-blue-400 hover:text-blue-600 hover:bg-blue-50 transition-all"
                    >
                      + Añadir día
                    </button>
                  )}
                </div>
              </div>
            );
          })}

          {/* Tarjeta para Añadir Nueva Semana */}
          {semanas.length < 4 && (
            <button 
              onClick={agregarSemana}
              className="bg-transparent border-2 border-dashed border-blue-300 rounded-2xl flex flex-col items-center justify-center p-8 text-blue-500 hover:bg-blue-50 hover:border-blue-400 transition-all min-h-[300px] h-full"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-12 w-12 mb-3 opacity-70" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v3m0 0v3m0-3h3m-3 0H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <span className="font-bold text-lg">Añadir nueva semana</span>
            </button>
          )}
        </div>
      </div>

      {/* Resumen Fijo (Footer) */}
      <div className="fixed bottom-0 left-0 right-0 bg-[#0f172a] text-white p-5 shadow-[0_-15px_30px_-10px_rgba(0,0,0,0.3)] pb-safe z-50">
        <div className="max-w-4xl mx-auto flex justify-between items-center">
          <div>
            <p className="text-[11px] uppercase tracking-wider text-gray-400 font-bold mb-1">Total ({semanas.length} Semanas)</p>
            <p className="font-medium text-lg text-gray-200">{formatoHorasMinutos(totalGeneral.horasDecimal)} <span className="text-xs text-gray-500 ml-1">hrs</span></p>
          </div>
          <div className="text-right">
            <p className="text-[11px] uppercase tracking-wider text-green-400 font-bold mb-1">Total a cobrar</p>
            <p className="text-3xl font-black text-white tracking-tighter">S/. {totalGeneral.pago.toFixed(2)}</p>
          </div>
        </div>
      </div>
    </main>
  )
}

export default App