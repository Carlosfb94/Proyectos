/**
 * Actualiza la tabla de stock con las compras recibidas.
 *
 * @param workbook - Libro activo en Excel Online.
 * @param stockTableName - Nombre de la tabla de stock (siempre "Stock").
 * @param purchases - Lista de compras. Cada subarreglo contiene:
 *   [0] Código, [1] Producto, [2] CantidadPedida, [3] FechaLlegada.
 * @returns Un objeto con la cantidad de códigos procesados.
 */
function main(
  workbook: ExcelScript.Workbook,
  stockTableName: string,
  purchases: (string | number)[][]
): { actualizados: number } {
  // Obtener la tabla de stock
  const tabla: ExcelScript.Table = workbook.getTable(stockTableName);
  const cuerpo: ExcelScript.Range = tabla.getRangeBetweenHeaderAndTotal();
  const filasExistentes: number = cuerpo.getRowCount();

  // Limpiar columnas D, F, G y H para eliminar pedidos antiguos
  if (filasExistentes > 0) {
    const vacio: string[][] = [];
    for (let i = 0; i < filasExistentes; i++) {
      vacio.push([""]);
    }
    // Columnas: D=3, F=5, G=6, H=7
    [3, 5, 6, 7].forEach((indice) => {
      tabla.getColumn(indice).getRangeBetweenHeaderAndTotal().setValues(vacio);
    });
  }

  // Valores actuales de la columna Código para localizar filas
  let codigos: string[][] = tabla.getColumn(0).getRangeBetweenHeaderAndTotal().getValues() as string[][];
  let procesados = 0;

  for (let compra of purchases) {
    const codigo: string = String(compra[0]);
    const producto: string = compra[1] ? String(compra[1]) : "";
    const cantidad: number = Number(compra[2]);
    const fechaLlegada: string = String(compra[3]);

    let fila: number = codigos.findIndex((c) => String(c[0]) === codigo);

    // Si no existe, insertar una nueva fila al final con Stock = 0
    if (fila === -1) {
      tabla.addRow(-1, [codigo, producto, 0, "", "", "", "", "", ""]);
      codigos.push([codigo]);
      fila = codigos.length - 1;
    }

    const filaRango: ExcelScript.Range = tabla.getRangeBetweenHeaderAndTotal().getRow(fila);
    const hoy: string = new Date().toISOString().substring(0, 10);

    // Actualizar datos de la fila
    filaRango.getCell(0, 3).setValue("Sí");
    filaRango.getCell(0, 5).setValue(hoy);
    filaRango.getCell(0, 6).setValue(cantidad);
    filaRango.getCell(0, 7).setValue(fechaLlegada);
    filaRango.getCell(0, 8).setFormula("=[@Stock]+[@CantidadPedida]");

    procesados++;
  }

  return { actualizados: procesados };
}

export default main;
