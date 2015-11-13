package unoil.examples;

import java.io.IOException;
import java.util.List;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.connection.NoConnectException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import ooo.connector.BootstrapSocketConnector;
import ooo.connector.server.OOoServer;

public class BootstrapSocketConnectorExample {

    private static final String OOO_EXEC_FOLDER = "C:/Program Files (x86)/OpenOffice 4/program";

    private static final String TEMPLATE_FOLDER = "C:/TEMP/";

    private static final String FILE_URL_PREFIX = "file:///";
    private static final String TEXT_DOCUMENT_EXTENSION = ".doc";
    private static final String PDF_DOCUMENT_EXTENSION = ".pdf";

    /**
     * Converts an OOo text document(C:/TEMP/OOo_doc.doc) to a PDF file using a
     * BootstrapConnector.
     * Be sure to have OpenOffice.org installed in order to execute this spike with success.
     * 
     * @param args The command line arguments
     */
    public static void main(String[] args) {
        String textDocumentName = "OOo_doc";

        String loadUrl = FILE_URL_PREFIX + TEMPLATE_FOLDER + textDocumentName + TEXT_DOCUMENT_EXTENSION;
        String storeUrl;

        try {
            storeUrl = FILE_URL_PREFIX + TEMPLATE_FOLDER + textDocumentName + "C" + PDF_DOCUMENT_EXTENSION;
            convertWithConnector(loadUrl, storeUrl);
        } catch (NoConnectException e) {
            System.out.println("OOo is not responding");
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.exit(0);
        }
    }

    private static void convertWithStaticConnector(String loadUrl, String storeUrl)
            throws Exception, IllegalArgumentException, IOException, BootstrapException {

        // Connect to OOo
        XComponentContext remoteContext = BootstrapSocketConnector.bootstrap(OOO_EXEC_FOLDER);

        // Convert text document to PDF
        convert(loadUrl, storeUrl, remoteContext);
    }

    private static void convertWithConnector(String loadUrl, String storeUrl)
            throws Exception, IllegalArgumentException, IOException, BootstrapException {

        // Create OOo server with additional -nofirststartwizard option
        List oooOptions = OOoServer.getDefaultOOoOptions();
        oooOptions.add("-nofirststartwizard");
        OOoServer oooServer = new OOoServer(OOO_EXEC_FOLDER, oooOptions);

        // Connect to OOo
        BootstrapSocketConnector bootstrapSocketConnector = new BootstrapSocketConnector(oooServer);
        XComponentContext remoteContext = bootstrapSocketConnector.connect();

        // Convert text document to PDF
        convert(loadUrl, storeUrl, remoteContext);

        // Disconnect and terminate OOo server
        bootstrapSocketConnector.disconnect();
    }

    protected static void convert(String loadUrl, String storeUrl, XComponentContext remoteContext)
            throws IllegalArgumentException, IOException, Exception {

        XComponentLoader xcomponentloader = getComponentLoader(remoteContext);

        Object objectDocumentToStore = xcomponentloader.loadComponentFromURL(loadUrl, "_blank", 0, new PropertyValue[0]);

        PropertyValue[] conversionProperties = new PropertyValue[1];
        conversionProperties[0] = new PropertyValue();
        conversionProperties[0].Name = "FilterName";
        conversionProperties[0].Value = "writer_pdf_Export";

        XStorable xstorable = UnoRuntime.queryInterface(XStorable.class, objectDocumentToStore);
        xstorable.storeToURL(storeUrl, conversionProperties);
    }

    private static XComponentLoader getComponentLoader(XComponentContext remoteContext) throws Exception {

        XMultiComponentFactory remoteServiceManager = remoteContext.getServiceManager();
        Object desktop = remoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", remoteContext);
        XComponentLoader xcomponentloader = UnoRuntime.queryInterface(XComponentLoader.class, desktop);

        return xcomponentloader;
    }

}
