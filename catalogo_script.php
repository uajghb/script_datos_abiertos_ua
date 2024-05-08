<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date as SharedDate;

date_default_timezone_set('Europe/Madrid');
$start = time();

# Carga del fichero .xlsx que contiene todos los metadatos (tabla de metadatos).
$filename = './La UA en cifras.xlsx';
$reader = IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(true);
$spreadsheet = $reader->load($filename);
$worksheet = $spreadsheet->getActiveSheet();
$worksheet_array = $worksheet->toArray();


# Primera fila como claves de cada columna.
$headings = array_shift($worksheet_array);
array_walk(
    $worksheet_array,
    function (&$row) use ($headings) {
        $row = array_combine($headings, $row);
    }
);


# Carga de un .xml base que sirve como punto de partida para la construcción del fichero final.
$content = file_get_contents('catalogo_base.rdf');
$dom = new DOMDocument();
$dom->preserveWhiteSpace = false;
$dom->formatOutput = true;
$dom->loadXML($content);
$root = $dom->documentElement;
$catalog=$root->getElementsByTagName('*')[0];

# Definición de variables auxiliares. 
$formatos = array('.pdf','.xlsx');
$formatos_mime = array(
    ".pdf" => "application/pdf",
    ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
);
$formatos_label = array(
    ".pdf" => "PDF",
    ".xlsx" => "XLSX" 
);

# Filtramos por formato, de este modo descartamos las entradas de la tabla que hacen referencia a páginas .html.
foreach($worksheet_array as $key=>$value){
    if (!in_array($value['EXTENSIÓN'], $formatos)){
        unset($worksheet_array[$key]);
    }
}
$worksheet_array = array_values($worksheet_array);

# A continuación, comienza la construcción del fichero .xml. 
# En primer lugar se añaden algunos metadatos globales del catálogo (fecha de creación y actualización).
# Nota : Se emplea el DateTime actual para la fecha de creación y de actualización del catálogo, puede que convenga
# establecer una fecha fija para la creación.

$date_now = new DateTime(date('m/d/Y H:i:s'));
$catalogIssued = $dom->createElement('dct:issued', $date_now->format(DateTime::ATOM));
$catalogIssued->setAttribute('rdf:datatype','http://www.w3.org/2001/XMLSchema#dateTime');
$catalog->appendChild($catalogIssued);
$catalogModified = $dom->createElement('dct:modified', $date_now->format(DateTime::ATOM));
$catalogModified->setAttribute('rdf:datatype','http://www.w3.org/2001/XMLSchema#dateTime');
$catalog->appendChild($catalogModified);

# Iteramos sobre cada fila de la tabla de metadatos, recogemos los metadatos de interés y 
# construimos etiquetas con esa información que se irán añadiendo al fichero base. Todo esto se realiza respetando una jerarquía de etiquetas 
# determinada.
foreach($worksheet_array as $row){
    $format = $row['EXTENSIÓN'];
    $dataset = $dom->createElement('dcat:dataset');
    $Dataset = $dom->createElement('dcat:Dataset');
    $distribution = $dom->createElement('dcat:distribution');
    $Distribution = $dom->createElement('dcat:Distribution');

    if ($format == '.xlsx'){
        $title_cas_value = getTitle($row['URL'], 'Castellano');
        $title_val_value = getTitle($row['URL'], 'Valenciano');
        $title_cas = $dom->createElement('dct:title', $title_cas_value);
        $title_val = $dom->createElement('dct:title', $title_val_value);
        $title_cas->setAttribute('xml:lang','es');
        $title_val->setAttribute('xml:lang','ca');
        $Dataset->appendChild($title_cas);
        $Dataset->appendChild($title_val);
        $Distribution->appendChild($title_cas->cloneNode(True));
        $Distribution->appendChild($title_val->cloneNode(True));

        $description_cas = $dom->createElement('dct:description', $title_cas_value);
        $description_val = $dom->createElement('dct:description', $title_val_value);
        $description_cas->setAttribute('xml:lang','es');
        $description_val->setAttribute('xml:lang','ca');
        $Dataset->appendChild($description_cas);
        $Dataset->appendChild($description_val);

        $language_es = $dom->createElement('dct:language','es');
        $language_ca = $dom->createElement('dct:language','ca');
        $Dataset->appendChild($language_es);
        $Dataset->appendChild($language_ca);


    } else{
        $title_value = $row['PÁGINA CONTENEDORA DEL ENLACE'];
        $title_value = $title_value . '. Año ' . explode('/', $row['URL'])[6];
        $title = $dom->createElement('dct:title', $title_value);
        $title->setAttribute('xml:lang','es');
        $Dataset->appendChild($title);
        $Distribution->appendChild($title->cloneNode(True));

        $description = $dom->createElement('dct:description', $title_value);
        $description->setAttribute('xml:lang','es');
        $Dataset->appendChild($description);

        $language_es = $dom->createElement('dct:language','es');
        $Dataset->appendChild($language_es);

    } 

    $theme = $dom->createElement('dcat:theme');
    $theme->setAttribute('rdf:resource', 'http://datos.gob.es/kos/sector-publico/sector/educacion');
    $Dataset->appendChild($theme);

    $publisher = $dom->createElement('dct:publisher');
    $publisher->setAttribute('rdf:resource','http://datos.gob.es/recurso/sector-publico/org/Organismo/U00100001');
    $Dataset->appendChild($publisher);

    $license = $dom->createElement('dct:license');
    $license->setAttribute('rdf:resource','https://si.ua.es/es/web-institucional-ua/normativa/copyright.html');
    $Dataset->appendChild($license);

    $spatial_comunidad = $dom->createElement('dct:spatial');
    $spatial_comunidad->setAttribute('rdf:resource', 'http://datos.gob.es/recurso/sector-publico/territorio/Autonomia/Comunitat-Valenciana');
    $spatial_pais = $dom->createElement('dct:spatial');
    $spatial_pais->setAttribute('rdf:resource', 'http://datos.gob.es/recurso/sector-publico/territorio/Pais/España');
    $Dataset->appendChild($spatial_comunidad);
    $Dataset->appendChild($spatial_pais);

    $issuedValue = $row['FECHA CREACIÓN'];
    $issued = $dom->createElement('dct:issued', SharedDate::excelToDateTimeObject($issuedValue)->format('c'));
    $issued->setAttribute('rdf:datatype','http://www.w3.org/2001/XMLSchema#dateTime');
    $Dataset->appendChild($issued);

    $modifiedValue = $row['FECHA MODIFICACIÓN'];
    $modified = $dom->createElement('dct:modified', SharedDate::excelToDateTimeObject($modifiedValue)->format('c'));
    $modified->setAttribute('rdf:datatype','http://www.w3.org/2001/XMLSchema#dateTime');
    $Dataset->appendChild($modified);

    $size = $dom->createElement('dcat:byteSize', $row['TAMAÑO']);
    $size->setAttribute('rdf:datatype', "http://www.w3.org/2001/XMLSchema#decimal");
    $Distribution->appendChild($size);

    $accessURL = $dom->createElement('dcat:accessURL', $row['URL']);
    $Distribution->appendChild($accessURL);

    $mediaType = $dom->createElement('dct:format');
    $imt = $mediaType->appendChild(
        $dom->createElement('dct:IMT')
    );
    $imt->appendChild(
        $dom->createElement('rdf:value', $formatos_mime[$format])
    );
    $imt->appendChild(
        $dom->createElement('rdfs:label', $formatos_label[$format])
    );
    $Distribution->appendChild($mediaType);


    $distribution->appendChild($Distribution);
    $Dataset->appendChild($distribution);
    $dataset->appendChild($Dataset);
    $catalog->appendChild($dataset);
	 
}

# Finalmente se guarda el fichero .xml generado con todos los metadatos, este fichero se empleará para federar el catálogo completo.

$dom->save('catalog.rdf');



# Función auxiliar para extraer el título de un conjunto de datos a partir de un fichero .xlsx.
# Dada la url del recurso y un idioma, se carga el fichero, se selecciona la hoja del fichero asociada a ese idioma y se extrae la información
# iterando sobre las filas. 
function getTitle($url, $idioma){
    try{
        # Se descarga el fichero y se guarda como "temp_file.xlsx".
        $file = file_get_contents($url);
        file_put_contents('temp_file.xlsx', $file);
        $reader = IOFactory::createReader('Xlsx');
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load('temp_file.xlsx');

        # Se carga la hoja del .xlsx asociada al idioma.
        # El nombre de las hojas no es consistente, se consideran todas las variaciones.
        $sheet_keys = array(
            "Castellano" => array('cas', 'formatos', 'formato cas', 'castellano', 'distribucion espacios cas', 'formatos'),
            "Valenciano" => array('val', 'formatos (val)', 'formato val', 'valenciano', 'distribucion espacios val', 'formats valencià') 
        );

        $sheet_names = $spreadsheet->getSheetNames();
        $worksheet = NULL;
        foreach($sheet_names as $name){
            if(in_array(strtolower(trim($name)),$sheet_keys[$idioma])){
                $worksheet = $spreadsheet->getSheetByName($name);
            }
        }

        # Se extrae el título.
        if (!is_null($worksheet)){
            $rows = $worksheet->toArray();
            $nonEmptyRows = array_filter($rows, 'checkRow');
            $nonEmptyRows = array_values($nonEmptyRows);
            
            # Se consideran las tres primeras filas no nulas.
            $title = $nonEmptyRows[0][0];
            for ($i=1;$i<3;$i++){
                $title .= ". " . $nonEmptyRows[$i][0];
            }
        } else{
            $title = 'No title';
        }
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        return $title;
    } catch (TypeError $e){
        return 'No title (error)';
    }
    

}

# Función auxiliar que se emplea en el método anterior para comprobar si una fila del fichero está vacía.
function checkRow($row){
    foreach($row as $column){
        if (!empty($column)){
            return true;
        }
    }
    return false;
}

$elapsed = time()-$start;
echo 'Saved ('.$elapsed.' s).';

?>