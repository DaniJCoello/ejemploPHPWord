<?php

namespace App\Http\Controllers;

use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Routing\Controller as BaseController;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\Writer\PDF;

class Controller extends BaseController
{
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    /**
     * Rellena los datos de un archivo .docx sin bloques
     */
    public function rellenarDatosSinBloques()
    {
        try {
            $rutaOriginal = 'prueba';
            $rutaDestino = 'prueba1';
            // Cambiar formato de doc a docx --> NO FUNCIONA
            // $phpWord = new PhpWordPhpWord();
            // $document = $phpWord->loadTemplate('anexo0.doc');
            // $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
            // $objWriter->save('anexo0.docx');

            // Se crea el objeto template con la ruta del archivo que se va a manipular
            // La ruta comienza en la carpeta public
            $template = new TemplateProcessor($rutaOriginal . '.docx');

            /*****************************************************************************************/
            /*******************ESTABLECER VALORES DE FORMA INDIVIDUAL EN EL DOCUMENTO****************/
            /*****************************************************************************************/
            // $template->setValue('nombre', 'Esther Colero');
            // $template->setValue('edad', rand(18,99));
            // $template->setValue('ciudad', 'Miguelturra');

            /*****************************************************************************************/
            /*********************************ESTABLECER VALORES EN MASA******************************/
            /*****************************************************************************************/
            // Los datos tienen que tener la misma estructura de índices que las variables
            $datos = [
                'nombre' => 'Johnny Melavo',
                'edad' => rand(18, 99),
                'ciudad' => 'Abenójar'
            ];

            // // Opción 1 --> Recorrerlos con un foreach
            // $variables = $template->getVariables();
            // // dd($variables);
            // foreach ($variables as $var) {
            //     $template->setValue($var, $datos[$var]);
            // }

            // // Opción 2 --> Establecerlos todos de una vez
            $template->setValues($datos);

            // Guardar el documento
            $template->saveAs($rutaDestino . '.docx');
            $this->convertirWordPDF($rutaDestino);
        } catch (Exception $e) {
        }
    }

    /**
     * Rellena los datos de un .docx usando bloques
     * No se puede hacer porque es una característica aún en desarrollo
     */
    public function rellenarBloques() {
        try {
            $template = new TemplateProcessor('bloques.docx');

            $template->saveAs('bloques1.docx');
        } catch (Exception $e) {

        }
    }

    /**
     * Esta función convierte un archivo word en pdf
     */
    private function convertirWordPDF(String $rutaArchivo)
    {
        /* Set the PDF Engine Renderer Path */
        $domPdfPath = base_path('vendor/dompdf/dompdf');
        \PhpOffice\PhpWord\Settings::setPdfRendererPath($domPdfPath);
        \PhpOffice\PhpWord\Settings::setPdfRendererName('DomPDF');

        // Load temporarily create word file
        $Content = \PhpOffice\PhpWord\IOFactory::load($rutaArchivo . '.docx');

        //Save it into PDF
        $savePdfPath = public_path($rutaArchivo. '.pdf');

        /*@ If already PDF exists then delete it */
        if ( file_exists($savePdfPath) ) {
            unlink($savePdfPath);
        }

        //Save it into PDF
        $PDFWriter = \PhpOffice\PhpWord\IOFactory::createWriter($Content,'PDF');
        $PDFWriter->save($savePdfPath);

        /*@ Remove temporarily created word file */
        if ( file_exists($rutaArchivo . '.docx') ) {
            unlink($rutaArchivo . '.docx');
        }
    }
}
