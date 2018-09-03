/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package MVC_CONTROLLER;

import com.jfoenix.controls.JFXButton;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.util.ResourceBundle;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javax.swing.JFileChooser;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * FXML Controller class
 *
 * @author cley
 */
public class MainController implements Initializable {

    @FXML
    private JFXButton bt_xml;
    private File file;

    List<String> data = new LinkedList<>();
    List<String> nf = new LinkedList<>();
    List<String> cod = new LinkedList<>();
    List<String> desc = new LinkedList<>();
    List<String> qtd = new LinkedList<>();
    List<String> valor = new LinkedList<>();
    List<String> venda = new LinkedList<>();
    List<String> ncm = new LinkedList<>();
    List<String> cest = new LinkedList<>();
    List<String> cst = new LinkedList<>();
    private String anf;
    private String adata;
    private String acst;

    /**
     * Initializes the controller class.
     */
    @Override
    public void initialize(URL url, ResourceBundle rb) {

        bt_xml.setOnMouseClicked((MouseEvent me) -> {
            pegarXML();
        });
    }

    public void pegarXML() {

        try {

            int cont = 0;
            //Cria um Arquivo Excel
            Workbook wb = new HSSFWorkbook();

            //Cria uma planilha Excel
            Sheet sheet = wb.createSheet("Dados do XML");

            //Cria uma linha na Planilha.
            Row cabecalho = sheet.createRow((short) 0);
            File arquivos[];
            FileChooser fileChooser1 = new FileChooser();
            fileChooser1.setTitle("Escolha um arquivo XML.");

            fileChooser1.setInitialDirectory(new File(System.getProperty("user.home")));

            file = fileChooser1.showOpenDialog(new Stage());

            File diretorio = new File(file.getParentFile() + "/");
            arquivos = diretorio.listFiles();
            System.out.println("arquivo " + arquivos.length);
            int rodan = 0;
            for (int a = 0; a < arquivos.length; a++) {
                //leia arquivos[i];

                //objetos para construir e fazer a leitura do documento
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                DocumentBuilder builder = factory.newDocumentBuilder();
                //abre e faz o parser de um documento xml de acordo com o nome passado no parametro
                Document doc = builder.parse(arquivos[a]);

                //Cria as células na linha
                cabecalho.createCell(0).setCellValue("Data");
                cabecalho.createCell(1).setCellValue("NECFE");
                cabecalho.createCell(2).setCellValue("CODIGO");
                cabecalho.createCell(3).setCellValue("DESCRIÇÃO");
                cabecalho.createCell(4).setCellValue("QTD");
                cabecalho.createCell(5).setCellValue("VALOR");
                cabecalho.createCell(6).setCellValue("VENDA");
                cabecalho.createCell(7).setCellValue("NCM");
                cabecalho.createCell(8).setCellValue("CEST");
                cabecalho.createCell(9).setCellValue("CST");
                System.out.println("criou a planilia");
                //Cria a segunda linha

                //cria uma lista de pessoas. Busca no documento todas as tag pessoa
                NodeList listaDePessoas = doc.getElementsByTagName("PISSN");

                //pego o tamanho da lista de pessoas
                int tamanhoLista = listaDePessoas.getLength();
                boolean valor = false;
                //varredura na lista de pessoas
                for (int i = 0; i < tamanhoLista; i++) {

                    //pego cada item (pessoa) como um nó (node)
                    Node noPessoa = listaDePessoas.item(i);
                    System.out.println("node de pessoa: " + noPessoa);
                    //verifica se o noPessoa é do tipo element (e não do tipo texto etc)
                    if (noPessoa.getNodeType() == Node.ELEMENT_NODE) {

                        //caso seja um element, converto o no Pessoa em Element pessoa
                        Element elementoPessoa = (Element) noPessoa;

                        //já posso pegar o atributo do element
                        String id = elementoPessoa.getAttribute("nItem");

                        //imprimindo o id
                        System.out.println("ID = " + id);

                        //recupero os nos filhos do elemento pessoa (nome, idade e peso)
                        NodeList listaDeFilhosDaPessoa = elementoPessoa.getChildNodes();

                        //pego o tamanho da lista de filhos do elemento pessoa
                        int tamanhoListaFilhos = listaDeFilhosDaPessoa.getLength();

                        //varredura na lista de filhos do elemento pessoa
                        for (int j = 0; j < tamanhoListaFilhos; j++) {

                            //crio um no com o cada tag filho dentro do no pessoa (tag nome, idade e peso)
                            Node noFilho = listaDeFilhosDaPessoa.item(j);

                            //verifico se são tipo element
                            if (noFilho.getNodeType() == Node.ELEMENT_NODE) {

                                //converto o no filho em element filho
                                Element elementoFilho = (Element) noFilho;

                                //verifico em qual filho estamos pela tag
                                switch (elementoFilho.getTagName()) {
                                    case "CST":
                                        //imprimo o nome
                                        System.out.println("CST=" + elementoFilho.getTextContent());
                                        cst.add(elementoFilho.getTextContent());
                                        acst = elementoFilho.getTextContent();
                                        break;

                                }
                            }
                        }
                    }
                }
                //cria uma lista de pessoas. Busca no documento todas as tag pessoa
                NodeList listaDeIDE = doc.getElementsByTagName("ide");

                //pego o tamanho da lista de pessoas
                int tamanhoListaIDE = listaDeIDE.getLength();
                for (int g = 0; g < tamanhoListaIDE; g++) {

                    //pego cada item (pessoa) como um nó (node)
                    Node noide = listaDeIDE.item(g);
                    System.out.println("node de pessoa: " + noide);
                    //verifica se o noPessoa é do tipo element (e não do tipo texto etc)
                    if (noide.getNodeType() == Node.ELEMENT_NODE) {

                        //caso seja um element, converto o no Pessoa em Element pessoa
                        Element elementoide = (Element) noide;

                        //recupero os nos filhos do elemento pessoa (nome, idade e peso)
                        NodeList listaDeFilhosDaide = elementoide.getChildNodes();

                        //pego o tamanho da lista de filhos do elemento pessoa
                        int tamanhoListaides = listaDeFilhosDaide.getLength();

                        //varredura na lista de filhos do elemento pessoa
                        for (int h = 0; h < tamanhoListaides; h++) {

                            //crio um no com o cada tag filho dentro do no pessoa (tag nome, idade e peso)
                            Node noFilhoide = listaDeFilhosDaide.item(h);

                            //verifico se são tipo element
                            if (noFilhoide.getNodeType() == Node.ELEMENT_NODE) {

                                //converto o no filho em element filho
                                Element elementoFilhoide = (Element) noFilhoide;

                                //verifico em qual filho estamos pela tag
                                switch (elementoFilhoide.getTagName()) {
                                    case "cNF":
                                        //imprimo o nome
                                        System.out.println("cNF=" + elementoFilhoide.getTextContent());
                                        nf.add(elementoFilhoide.getTextContent());
                                        anf = elementoFilhoide.getTextContent();
                                        break;
                                    case "dEmi":
                                        System.out.println("dEmi=" + elementoFilhoide.getTextContent());
                                        data.add(elementoFilhoide.getTextContent());
                                        adata = elementoFilhoide.getTextContent();
                                        break;
                                }
                            }
                        }
                    }
                }
                //cria uma lista de pessoas. Busca no documento todas as tag pessoa
                NodeList listaDeprod = doc.getElementsByTagName("prod");

                //pego o tamanho da lista de pessoas
                int tamanhoListaprod = listaDeprod.getLength();
                System.err.println("Tamanho produto:        " + tamanhoListaprod);
                for (int g = 0; g < tamanhoListaprod; g++) {

                    //pego cada item (pessoa) como um nó (node)
                    Node noprod = listaDeprod.item(g);
                    //verifica se o noPessoa é do tipo element (e não do tipo texto etc)
                    if (noprod.getNodeType() == Node.ELEMENT_NODE) {

                        //caso seja um element, converto o no Pessoa em Element pessoa
                        Element elementoprod = (Element) noprod;

                        //recupero os nos filhos do elemento pessoa (nome, idade e peso)
                        NodeList listaDeFilhosDaprod = elementoprod.getChildNodes();

                        //pego o tamanho da lista de filhos do elemento pessoa
                        int tamanhoListaprods = listaDeFilhosDaprod.getLength();

                        //varredura na lista de filhos do elemento pessoa
                        System.err.println("Tamanho de produtos:    " + tamanhoListaprods);
                        for (int h = 0; h < tamanhoListaprods; h++) {

                            //crio um no com o cada tag filho dentro do no pessoa (tag nome, idade e peso)
                            Node noFilhoprod = listaDeFilhosDaprod.item(h);

                            //verifico se são tipo element
                            if (noFilhoprod.getNodeType() == Node.ELEMENT_NODE) {

                                //converto o no filho em element filho
                                Element elementoFilhoprod = (Element) noFilhoprod;

                                //verifico em qual filho estamos pela tag
                                switch (elementoFilhoprod.getTagName()) {
                                    case "cProd":
                                        //imprimo o nome
                                        System.out.println("cProd=" + elementoFilhoprod.getTextContent());
                                        cod.add(elementoFilhoprod.getTextContent());
                                        cont++;
                                        rodan++;
                                        break;
                                    case "xProd":
                                        System.out.println("dEmi=" + elementoFilhoprod.getTextContent());
                                        desc.add(elementoFilhoprod.getTextContent());
                                        break;
                                    case "NCM":
                                        //imprimo o nome
                                        System.out.println("NCM=" + elementoFilhoprod.getTextContent());
                                        ncm.add(elementoFilhoprod.getTextContent());
                                        break;
                                    case "qCom":
                                        System.out.println("qCom=" + elementoFilhoprod.getTextContent());
                                        qtd.add(elementoFilhoprod.getTextContent());
                                        break;
                                    case "vUnCom":
                                        //imprimo o nome
                                        System.out.println("vUnCom=" + elementoFilhoprod.getTextContent());
                                        this.valor.add(elementoFilhoprod.getTextContent());
                                        break;
                                    case "vProd":
                                        System.out.println("vProd=" + elementoFilhoprod.getTextContent());
                                        venda.add(elementoFilhoprod.getTextContent());
                                        break;

                                }

                            }
                        }

                    }

                }
                //cria uma lista de pessoas. Busca no documento todas as tag pessoa
                NodeList listaDecest = doc.getElementsByTagName("obsFiscoDet");

                //pego o tamanho da lista de pessoas
                int tamanhoListacest = listaDecest.getLength();
                for (int g = 0; g < tamanhoListacest; g++) {

                    //pego cada item (pessoa) como um nó (node)
                    Node nocest = listaDecest.item(g);
                    //verifica se o noPessoa é do tipo element (e não do tipo texto etc)
                    if (nocest.getNodeType() == Node.ELEMENT_NODE) {

                        //caso seja um element, converto o no Pessoa em Element pessoa
                        Element elementocest = (Element) nocest;

                        //recupero os nos filhos do elemento pessoa (nome, idade e peso)
                        NodeList listaDeFilhosDacest = elementocest.getChildNodes();

                        //pego o tamanho da lista de filhos do elemento pessoa
                        int tamanhoListacests = listaDeFilhosDacest.getLength();

                        //varredura na lista de filhos do elemento pessoa
                        for (int h = 0; h < tamanhoListacests; h++) {

                            //crio um no com o cada tag filho dentro do no pessoa (tag nome, idade e peso)
                            Node noFilhocest = listaDeFilhosDacest.item(h);

                            //verifico se são tipo element
                            if (noFilhocest.getNodeType() == Node.ELEMENT_NODE) {

                                //converto o no filho em element filho
                                Element elementoFilhocest = (Element) noFilhocest;

                                //verifico em qual filho estamos pela tag
                                switch (elementoFilhocest.getTagName()) {
                                    case "xTextoDet":
                                        //imprimo o nome
                                        System.out.println("cest=" + elementoFilhocest.getTextContent());
                                        cest.add(elementoFilhocest.getTextContent());
                                        break;
                                }
                            }
                        }
                    }
                }

                for (int q = 0; q < (tamanhoListaprod - 1); q++) {

                    cst.add(acst);
                    data.add(adata);
                    nf.add(anf);

                }

            }
            System.out.println("cont" + cont);
            for (int d = 0; d < rodan; d++) {
                Row dados = sheet.createRow((short) (d + 1));

                //Nas células a seguir vc substitui pelos valores das notas
                dados.createCell(0).setCellValue(data.get(d));
                dados.createCell(1).setCellValue(nf.get(d));
                dados.createCell(2).setCellValue(cod.get(d));
                dados.createCell(3).setCellValue(desc.get(d));
                dados.createCell(4).setCellValue(qtd.get(d));
                dados.createCell(5).setCellValue(this.valor.get(d));
                dados.createCell(6).setCellValue(venda.get(d));
                dados.createCell(7).setCellValue(ncm.get(d));
                dados.createCell(8).setCellValue(cest.get(d));
                dados.createCell(9).setCellValue(cst.get(d));
            }
            FileChooser chooser = new FileChooser();
            chooser.setTitle("Escolha o nome da planilia e onde ira salva-lá.");
            System.out.println("total:  " + rodan);

            String caminho = null;
            File retorno = chooser.showSaveDialog(new Stage());
            if (retorno != null) {
                System.out.println("caminho:    " + retorno.getAbsolutePath());
                caminho = retorno.getAbsolutePath() + ".xlsx";

                try (FileOutputStream fileOut = new FileOutputStream(caminho)) {
                    wb.write(fileOut);
                }
            }
        } catch (ParserConfigurationException | SAXException | IOException ex) {
            Logger.getLogger(MainController.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
