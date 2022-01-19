<?php

namespace App\Http\Controllers;

use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Routing\Controller as BaseController;
use PhpOffice\PhpWord;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord as PhpWordPhpWord;
use PhpOffice\PhpWord\TemplateProcessor;

class Controller extends BaseController
{
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    /**
     * Rellena los datos de un archivo .docx sin bloques
     */
    public function rellenarDatosSinBloques()
    {
        try {
            // Cambiar formato de doc a docx --> NO FUNCIONA
            // $phpWord = new PhpWordPhpWord();
            // $document = $phpWord->loadTemplate('anexo0.doc');
            // $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
            // $objWriter->save('anexo0.docx');

            // Se crea el objeto template con la ruta del archivo que se va a manipular
            // La ruta comienza en la carpeta public
            $template = new TemplateProcessor('prueba.docx');

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
            $template->saveAs('prueba1.docx');
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
}
