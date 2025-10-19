package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

public class GenericFactorPanel extends JPanel {
    private final ResourceBundle messages;
    private JLabel fileLabel;
    private JComboBox<String> sheetSelector;
    private JButton addFileButton;
    private JButton applyAndSaveButton;
    private JTable previewTable;
    private JTable factorsTable;

    public GenericFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout(5,5));
        setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel filePanel = new JPanel(new BorderLayout()); filePanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.file.management"))); filePanel.setBackground(UIUtils.CONTENT_BACKGROUND); filePanel.setPreferredSize(new Dimension(300,200));
        JPanel contentPanel = new JPanel(); contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS)); contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addFileButton = new JButton(messages.getString("button.file.add")); addFileButton.setAlignmentX(Component.CENTER_ALIGNMENT); UIUtils.styleButton(addFileButton);
        fileLabel = new JLabel(messages.getString("label.file.none")); fileLabel.setForeground(Color.GRAY); fileLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.select")); sheetLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        sheetSelector = new JComboBox<>(); sheetSelector.setAlignmentX(Component.CENTER_ALIGNMENT); sheetSelector.setMaximumSize(new Dimension(150,25)); UIUtils.styleComboBox(sheetSelector);
        contentPanel.add(Box.createVerticalStrut(5)); contentPanel.add(addFileButton); contentPanel.add(Box.createVerticalStrut(5)); contentPanel.add(fileLabel); contentPanel.add(Box.createVerticalStrut(5)); contentPanel.add(sheetLabel); contentPanel.add(Box.createVerticalStrut(5)); contentPanel.add(sheetSelector);
        filePanel.add(contentPanel, BorderLayout.CENTER);
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER)); buttonPanel.setBackground(UIUtils.CONTENT_BACKGROUND); applyAndSaveButton = new JButton(messages.getString("button.apply.save.excel")); applyAndSaveButton.setEnabled(false); UIUtils.styleButton(applyAndSaveButton); buttonPanel.add(applyAndSaveButton); filePanel.add(buttonPanel, BorderLayout.SOUTH);

        // Data management panel
        JPanel dataPanel = new JPanel(new BorderLayout()); dataPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.data.management"))); dataPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        String[] columnNames = { messages.getString("table.header.entity"), messages.getString("table.header.year"), messages.getString("table.header.factor"), messages.getString("table.header.unit") };
        DefaultTableModel model = new DefaultTableModel(columnNames, 0) { @Override public boolean isCellEditable(int row, int column) { return false; } };
        factorsTable = new JTable(model); factorsTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION); factorsTable.getTableHeader().setReorderingAllowed(false); UIUtils.styleTable(factorsTable);
        JScrollPane scrollPane = new JScrollPane(factorsTable); scrollPane.setBorder(BorderFactory.createEmptyBorder(0,0,0,0)); dataPanel.add(scrollPane, BorderLayout.CENTER);
        JPanel buttonPanel2 = new JPanel(new FlowLayout(FlowLayout.RIGHT)); buttonPanel2.setBackground(UIUtils.CONTENT_BACKGROUND);
        JButton addButton = new JButton(messages.getString("button.add")); JButton editButton = new JButton(messages.getString("button.edit")); JButton deleteButton = new JButton(messages.getString("button.delete")); UIUtils.styleButton(addButton); UIUtils.styleButton(editButton); UIUtils.styleButton(deleteButton); buttonPanel2.add(addButton); buttonPanel2.add(editButton); buttonPanel2.add(deleteButton); dataPanel.add(buttonPanel2, BorderLayout.SOUTH);

        // Preview panel
        JPanel previewPanel = new JPanel(new BorderLayout()); previewPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.preview"))); previewPanel.setBackground(UIUtils.CONTENT_BACKGROUND); previewPanel.setPreferredSize(new Dimension(0,350));
        previewTable = new JTable(); previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF); previewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION); UIUtils.styleTable(previewTable);
        JScrollPane previewScrollPane = new JScrollPane(previewTable); previewScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); previewScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS); UIUtils.setupPreviewTable(previewTable); previewPanel.add(previewScrollPane, BorderLayout.CENTER);

        setLayout(new BorderLayout()); add(filePanel, BorderLayout.NORTH); add(dataPanel, BorderLayout.CENTER); add(previewPanel, BorderLayout.SOUTH);
    }

    public JLabel getFileLabel() { return fileLabel; }
    public JComboBox<String> getSheetSelector() { return sheetSelector; }
    public JButton getApplyAndSaveButton() { return applyAndSaveButton; }
    public JTable getPreviewTable() { return previewTable; }
    public JTable getFactorsTable() { return factorsTable; }
}
