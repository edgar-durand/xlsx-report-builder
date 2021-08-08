const xlsx = require('xlsx');
const nodeXlsx = require('node-xlsx');
const readdirp = require('readdirp');
const fs = require('fs');

const SETIINGS = {
    root: './path/',
    entryType: 'all',
    fileFilter: ['*.xlsx', '*.xls'],
    // directoryFilter: [],
    // type: 'files' | 'directories' | 'files_directories' | 'all',
    // depth: number,
};

// En este ejemplo, esta variable almacenará todas las rutas de los archivos y directorios dentro de la ruta proporcionada
const allFilePaths = [];
const names = [];
const totalBus = [];
const PNA_BOOK = [
    ['MATRICULA',
        'CLASIFICACION',
        'DESDE',
        'HASTA',
        'TIEMPO',
        'LUGAR',
        'DESCRIPCION'
    ],
];
let totalTracks = 0;

/************************************ */

// AQUI LAS PARADAS CONOCIDAS COMO AUTORIZADAS SE DEBEN REGISTRAR
// AUTORIZOS DE PARQUEOS, PARQUEOS EVENTUALES, RUTAS AUTORIZAS ETC.
// LAS QUE SE ENCUENTREN EN ESTE REGISTRO NO APARECERAN EN EL REPORTE <PNA>
/************************************ */
const ALLOWED_STOP = [
    // KM 0 Zona recojida de los trabajadores
    'PARQUE LA GUIRA, BANES',
    'Antonio Dumois Entre Pasaje 8 y Coronel Tío, La Palma Número Uno, Banes, Holguín', //Entrada de la base
    'Guamá Entre Carretera de Veguita y Final, Veguitas, Banes, Holguín', //Entrada al cupet

    'Flor Crombet Entre Bayamo y Carlos Manuel de Céspedes, Banes, Banes, Holguín', //Parque Cardenas
    'Ave.General Marrero Entre Bayamo y Carlos Manuel de Céspedes, Banes, Banes, Holguín', //Parque Cardenas
    'Luz y Caballero Entre Flor Crombet y Augusto Blanco, Banes, Banes, Holguín', //Parque Cardenas
    'Augusto Blanco Entre Bruno Meriño y Bayamo, Banes, Banes, Holguín', //Parque Cardenas
    'Ave.Cárdenas Entre Carlos Manuel de Céspedes y José Martí, Banes, Banes, Holguín', //Parque Cardenas

    // EN RUTAS ACTIVIDAD FUNDAMENTAL
    'El Limpio de Retrete,  BANES,  HOLGUÍN',
    'TERMINAL ÓMNIBUS BANES, BANES',
    'Playa Pesquero,  BANES,  HOLGUÍN',
    'HOTEL BRISAS GUARDALAVACA, GUARDALAVACA',
    'HOTEL PLAYA PESQUERO, PESQUERO',
    'ÓPTICA, BANES',
    'Playa Pesquera Nueva,  RAFAEL FREYRE,  HOLGUÍN',
    'Melilla,  RAFAEL FREYRE,  HOLGUÍN',
    'Tonquín,  RAFAEL FREYRE,  HOLGUÍN',
    'Guardalavaca,  BANES,  HOLGUÍN',
    'HOTEL GUARDALAVACA, GUARDALAVACA',
    'Playa de Morales,  -,  HOLGUÍN',
    'DESTIERRO / ACUEDUCTO, GUARDALAVACA',
    'HOTEL RIO DE ORO, BAHÍA DE NARANJO',
    'CONUCO MONGO VIÑA, BAHÍA DE NARANJO',
    'HOTEL RÍO DE LUNAS, BAHÍA DE NARANJO',
    'El Ramón,  ANTILLA,  HOLGUÍN',


    // EN VIAJES EVENTUALES...
    // HOLGUIN...
    'Loma la Cruz,  HOLGUÍN,  HOLGUÍN',
    'islazul pernik,  HOLGUÍN,  HOLGUÍN',

    // MOA
    'Moa,  MOA,  HOLGUÍN',
];

/************************************ */

// AQUI LAS PARADAS CONOCIDAS COMO NO AUTORIZADAS
/************************************ */
const FORBIDDEN_STOP = [
    //EX. ['TRAFFIC STREET, NO. 18', 'CHOFER'S HOUSE'],
    ['Los Pasos,  BANES,  HOLGUÍN', 'CASA DE ALEXIS RODRIGO JIMENEZ'],
    ['HOSPITAL CALLE MULA, BANES', 'RECOJIDA DEL CHOFER LUIS O. RICARDO AVILA'],
];

/************************************ */

// Iterar recursivamente a través de una carpeta
readdirp(SETIINGS.root, {
        ...SETIINGS
    })
    .on('data', function (entry) {
        // ejecutar cada vez que se encuentre un archivo en el directorio de proveedores
        // Almacene la ruta completa del archivo / directorio en nuestra matriz personalizada
        allFilePaths.push(
            entry.fullPath
        );
        names.push(
            entry.path
        )
        totalTracks++;
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
            const workbook = xlsx.readFile(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let sheet = xlsx.utils.sheet_to_json(worksheet, {
                header: 1
            });

            // Delete the 7 first rows in sheet
            for (let index = 0; index < 7; index++) {
                delete sheet[index];
            }

            const book = sheet[7];
            const busTag = book[0];
            let flag = 0;

            return sheet.map((row) => {

                // AQUI AGREGAS LAS JUSTIFICACIONES DE LAS PARADAS
                /************************************ */
                const ANALIZE = [
                    ['TRANSMETRO BANES, CORONEL TÍO RPTO TORRENTERAS BANES', 'PARQUEO'],
                    ['La Palma Número Uno,  BANES,  HOLGUÍN', 'PARQUEO'],
                    ['CUPET VEGUITA, VEGUITAS BANES HOLG', 'HABILITANDO'],
                    ['Carretera de Veguita Entre Guamá y Camino al Vivero, Veguitas, Banes, Holguín', 'HABILITANDO'],
                ];
                /************************************ */

                row[1] = 'Detenciones';
                row[6] = '';

                // ACA SE JUSTIFICAN LAS PARADAS CON LA RELACION PROVISTA EN LA CONSTANTE <ANALIZE>
                ANALIZE.forEach((couple) => {
                    ALLOWED_STOP.push(couple[0]);
                    if (row[5] === couple[0]) row[6] = couple[1];
                });
                /************************************************ */

                // ACA LLENAMOS LA RELACION DE LAS PARADAS NO AUTORIZADAS QUE SON MAYORES A 5 MIN
                if (!ALLOWED_STOP.includes(row[5])) { //SI NO ESTA INCLUIDA EN LA LISTA DE PARADAS AUTORIZADAS
                    if (+String(row[4]).toString().substr(0, 2) > 0 || +String(row[4]).toString().substr(3, 2) > 5) { // SI ES MAYOR DE 5 MIN
                        const JUSTIFY = Array.from(row);

                        if (flag === 0)
                            JUSTIFY[0] = busTag;
                        flag++;
                        JUSTIFY[1] = 'PNA';
                        JUSTIFY[6] = '';
                        //  NOW WE TRY TO JUSTIFY WITH KNOWN PROVIDED FORBIDDEN STOP
                        FORBIDDEN_STOP.forEach((couples) => {
                            if (JUSTIFY[5] === couples[0]) JUSTIFY[6] = couples[1];
                        });
                        PNA_BOOK.push(JUSTIFY);
                    }
                }

                /********************************************************** */

                row[6] = row[6] === '' ? 'TRASLADO DE PERSONAL' : row[6];
                return row;
            });
        }

        const sorted = [];

        // ACA ARMAMOS EL LIBRO CON LA ESTRUCTURA DE UN ARRAY DE FILAS
        allFilePaths.forEach(book => {
            const sheet = handleLoad(book);
            for (let index = 0; index < 7; index++) {
                sheet.shift();
            }
            sheet.forEach((row) => sorted.push(row));
        });


        // const range = {s: {c: 0, r:0 }, e: {c:0, r:3}}; // A1:A4
        const options = {
            '!cols': [{
                wch: 20
            }, {
                wch: 15
            }, {
                wch: 20
            }, {
                wch: 20
            }, {
                wch: 10
            }, {
                wch: 45
            }, {
                wch: 40
            }],
            // '!merges': [ range ],
        };

        const date = new Date();
        const day = date.getDate().toString().length > 1 ? date.getDate() : `0${date.getDate()}`;
        const month = date.getMonth().toString().length > 1 ? date.getMonth() + 1 : `0${date.getMonth() + 1}`;
        const year = date.getFullYear();

        const buffer = nodeXlsx.build([{
            name: `report`,
            data: sorted
        }, {
            name: `PNA-total-${PNA_BOOK.length-1}`,
            data: PNA_BOOK
        }], options);
        fs.writeFileSync(`./Report-${year}${month}${day}--tracks-${totalTracks}-bus-${totalBus.length}.xlsx`, buffer, 'binary');
    });