package org.example;
import javax.swing.*;
import java.awt.*;
import java.io.File;

public class MainUI {

        public static void main(String[] args) {

            JFrame frame = new JFrame("Tax System 💰");
            frame.setSize(400, 250);
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setLayout(new FlowLayout());

            JButton fileButton = new JButton("Choose Excel");
            JLabel fileLabel = new JLabel("No file selected");

            final File[] selectedFile = new File[1];

            fileButton.addActionListener(e -> {
                JFileChooser chooser = new JFileChooser();
                int result = chooser.showOpenDialog(null);
                if (result == JFileChooser.APPROVE_OPTION) {
                    selectedFile[0] = chooser.getSelectedFile();
                    fileLabel.setText(selectedFile[0].getName());
                }
            });

            String[] months = {
                    "1 - Jan","2 - Feb","3 - Mar","4 - Apr",
                    "5 - May","6 - Jun","7 - Jul","8 - Aug",
                    "9 - Sep","10 - Oct","11 - Nov","12 - Dec"
            };

            JComboBox<String> monthBox = new JComboBox<>(months);

            JButton runButton = new JButton("Run");
            JLabel resultLabel = new JLabel("");

            runButton.addActionListener(e -> {

                if (selectedFile[0] == null) {
                    resultLabel.setText("❌ Choose file");
                    return;
                }

                int month = monthBox.getSelectedIndex() + 1;

                try {
                    Main.runTax(selectedFile[0].getPath(), month);
                    resultLabel.setText("✅ Done Month " + month);
                } catch (Exception ex) {
                    ex.printStackTrace();
                    resultLabel.setText("❌ Error");
                }
            });

            frame.add(fileButton);
            frame.add(fileLabel);
            frame.add(monthBox);
            frame.add(runButton);
            frame.add(resultLabel);

            frame.setVisible(true);
        }
}
