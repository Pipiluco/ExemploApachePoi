package br.com.lucasfrancisco.exemploapachepoi;

import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.provider.OpenableColumns;
import android.support.annotation.Nullable;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class MainActivity extends AppCompatActivity {
    private TextView tvResultado;
    private Button btnLerArquivo, btnPathArquivo;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        tvResultado = (TextView) findViewById(R.id.tvResultado);
        btnLerArquivo = (Button) findViewById(R.id.btnLerArquivo);
        btnPathArquivo = (Button) findViewById(R.id.btnPathArquivo);

        btnLerArquivo.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                //lerArquivo();
                Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
                intent.addCategory(Intent.CATEGORY_OPENABLE);
                intent.setType("*/xlsx");
                intent.putExtra(Intent.EXTRA_TITLE, "Teste");
                startActivityForResult(intent, 2);
            }
        });

        btnPathArquivo.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intent = new Intent();
                intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true);
                intent.setAction(Intent.ACTION_GET_CONTENT);
                startActivityForResult(Intent.createChooser(intent, "Arquivos"), 1);
            }
        });
    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (resultCode == RESULT_OK) {
            switch (requestCode) {
                case 1:
                    requestLoadFile(data);
                    break;
                case 2:
                    // requestCreateFile(data);
                    tvResultado.setText(data.getData().toString());

                    break;
            }
        }
    }

    // OK
    public void lerArquivo() {
        try {
            String nomeArquivo = "teste.xls";
            File arquivo = new File(Environment.getExternalStorageDirectory().getAbsolutePath(), "/Download/" + nomeArquivo); //  /storage/emulated/0/Download/teste.xls

            Log.e("PATH", arquivo.toString());

            FileInputStream fileInputStream = new FileInputStream(arquivo);


//
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileInputStream);
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
            Cell cell = hssfSheet.getRow(1).getCell(1);
            String valor = cell.getStringCellValue();

            tvResultado.setText(valor);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void requestLoadFile(Intent intent) {
        if (intent.getClipData() != null) {
            int totalArquivosSelecionado = intent.getClipData().getItemCount();

            for (int i = 0; i < totalArquivosSelecionado; i++) {
                Uri uri = intent.getClipData().getItemAt(i).getUri();
                String nomeArquivo = uri + getNomeArquivo(uri);

                tvResultado.append(nomeArquivo + "\n");
            }
        } else if (intent.getData() != null) {
            Uri uri = intent.getData();

            try {
                InputStream inputStream = getApplicationContext().getContentResolver().openInputStream(uri);

                Log.d("INPUT", inputStream.toString());

                XSSFWorkbook hssfWorkbook = new XSSFWorkbook(inputStream);
                XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
                Cell cell = hssfSheet.getRow(1).getCell(1);
                String valor = "";

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        valor = String.valueOf(cell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_STRING:
                        valor = String.valueOf(cell.getStringCellValue());
                        break;
                }

                tvResultado.setText(valor);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void requestCreateFile(Intent intent) {
        if (intent.getData() != null) {
            try {
                OutputStream outputStream = getApplicationContext().getContentResolver().openOutputStream(intent.getData());

                Log.d("INPUT", outputStream.toString());

                ManagerXLSX managerXLSX = new ManagerXLSX();
                managerXLSX.criar(outputStream);

                tvResultado.setText(intent.getData().getPath());

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }
    }

    // Retorna o nome de um arquivo selecionado no gerenciador de arquivos
    public String getNomeArquivo(Uri uri) {
        String resultado = null;

        if (uri.getScheme().equals("content")) {
            Cursor cursor = getContentResolver().query(uri, null, null, null, null);
            try {
                if (cursor != null && cursor.moveToFirst()) {
                    resultado = cursor.getString(cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME));
                }
            } finally {
                cursor.close();
            }
        }

        if (resultado == null) {
            resultado = uri.getPath();
            int corte = resultado.lastIndexOf('/');
            if (corte != -1) {
                resultado = resultado.substring(corte + 1);
            }
        }
        return resultado;
    }
}
