const xlsx = require('xlsx');
const nodeXlsx = require('node-xlsx');
const readdirp = require('readdirp');
const fs = require('fs');

const SETIINGS = {
    root: './path/',
    entryType: 'all',
    // fileFilter: [],
    // directoryFilter: [],
    // type: 'files' | 'directories' | 'files_directories' | 'all',
    // depth: number,
};

// En este ejemplo, esta variable almacenará todas las rutas de los archivos y directorios dentro de la ruta proporcionada
const allFilePaths = [];
const names = [];
let totalTracks = 0;
let totalBus = [];

// Iterar recursivamente a través de una carpeta
readdirp(SETIINGS.root, { ...SETIINGS })
    .on('data', function (entry) {
        // ejecutar cada vez que se encuentre un archivo en el directorio de proveedores
        // Almacene la ruta completa del archivo / directorio en nuestra matriz personalizada
        allFilePaths.push(
            entry.fullPath
        );
        names.push(
            entry.path
        )
        totalTracks ++;
    })
    .on('warn', function (warn) {
        console.log("Warn: ", warn);
    })
    .on('error', function (err) {
        console.log("Error: ", err);
    })
    .on('end', function () {

        names.forEach((name) => {
            if (!(totalBus.includes(String(name).substr(0, 7))))
                totalBus.push(String(name).substr(0, 7))
        })

        /**
         * 
         * @param {Buffer} file 
         * @returns 
         */
        const handleLoad = (file) => {
            let workbook = xlsx.readFile(file);
            let worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let sheet = xlsx.utils.sheet_to_json(worksheet, {
                header: 1
            });
            for (let index = 0; index < 7; index++) {
                delete sheet[index];                
            }
            
            return sheet.map((row) => {
                // AQUI AGREGAS LAS JUSTIFICACIONES DE LAS PARADAS
                const ANALIZE = [
                    ['TRANSMETRO BANES, CORONEL TÍO RPTO TORRENTERAS BANES', 'PARQUEO'],
                    ['La Palma Número Uno,  BANES,  HOLGUÍN', 'PARQUEO'],
                    ['CUPET VEGUITA, VEGUITAS BANES HOLG', 'HABILITANDO'],
                    ['Carretera de Veguita Entre Guamá y Camino al Vivero, Veguitas, Banes, Holguín', 'HABILITANDO'],
                ]; 
                row[1] = 'Detenciones';
                row[6] = '';
                ANALIZE.forEach((couple) => {
                    if (row[5] === couple[0]) row[6] = couple[1];
                });

                row[6] = row[6] === '' ? 'TRASLADO DE PERSONAL' : row[6];
                return row;
            });
        }

        const sorted = [];

        allFilePaths.forEach(book => {
            const sheet = handleLoad(book);
            for (let index = 0; index < 7; index++) {
                sheet.shift();                
            }
            sheet.forEach((row) => sorted.push(row));          
        });
        // const range = {s: {c: 0, r:0 }, e: {c:0, r:3}}; // A1:A4
        const options = {
            '!cols': [{ wch:20 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 10 }, { wch: 45 }, { wch: 20 } ],
            // '!merges': [ range ],
        };
        
        const date = new Date();
        const day = date.getDate().toString().length > 1 ? date.getDate() : `0${date.getDate()}`;
        const month = date.getMonth().toString().length > 1 ? date.getMonth() + 1 : `0${date.getMonth() + 1}`;
        const year = date.getFullYear();

        const buffer = nodeXlsx.build([{ name: `report`, data: sorted }], options);
        fs.writeFileSync(`./Report-${year}${month}${day}--tracks-${totalTracks}-bus-${totalBus.length}.xlsx`, buffer, 'binary')
    });