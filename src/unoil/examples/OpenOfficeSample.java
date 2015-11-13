package unoil.examples;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import ooo.connector.BootstrapSocketConnector;

public class OpenOfficeSample {

    public static void main(String[] args) {

        /*
         * First thing to do is to connect to the OpenOffice. 
         * The following code will start an instance of OpenOffice.
         */
        XComponentContext xContext = null;
        Object oDesktop = null;
        String oooExeFolder = "C:/Program Files (x86)/OpenOffice 4/program";
        try {
            // Get the remote office component context
//            xContext = Bootstrap.bootstrap();
            xContext = BootstrapSocketConnector.bootstrap(oooExeFolder);
            // Get the remote office service manager
            XMultiComponentFactory xMCF = xContext.getServiceManager();
            // Get the root frame (i.e. desktop) of openoffice framework.
            oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
        } catch (Exception e) {
            e.printStackTrace();
        }
        // Desktop has 3 interfaces. The XComponentLoader interface provides ability to load components.
        XComponentLoader xCLoader = UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);

        /*
         * Now, we have to create an empty document in the new writer window.
         */
        // Create a document
        XComponent document = null;
        try {
            document = xCLoader.loadComponentFromURL("private:factory/swriter", "_blank", 0, new PropertyValue[0]);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IllegalArgumentException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        // Get the textdocument
        XTextDocument aTextDocument = UnoRuntime.queryInterface(com.sun.star.text.XTextDocument.class, document);

        // Get its text
        XText xText = aTextDocument.getText();

        /*
         * Now i can append my first piece text to the document.
         */
        // Adding text to document
        xText.insertString(xText.getEnd(), "My First OpenOffice Document", false);

        /*
         * The final task is to save the document. Here is the code to do that
         */
        // the url where the document is to be saved
        String storeUrl = "file:///C:/TEMP/OOo_doc.odt";

        // Save the document
        XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, document);
        PropertyValue[] storeProps = new PropertyValue[0];
        try {
            xStorable.storeAsURL(storeUrl, storeProps);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

}
