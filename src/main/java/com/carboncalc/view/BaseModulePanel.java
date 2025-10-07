package com.carboncalc.view;

import com.carboncalc.util.UIUtils;
import javax.swing.*;
import java.awt.*;
import java.util.ResourceBundle;

public abstract class BaseModulePanel extends JPanel {
    protected final ResourceBundle messages;
    protected final JPanel contentPanel;
    
    public BaseModulePanel(ResourceBundle messages) {
        this.messages = messages;
        this.contentPanel = new JPanel(new BorderLayout());
        setLayout(new BorderLayout());
        add(contentPanel, BorderLayout.CENTER);
        UIUtils.stylePanel(this);
        UIUtils.stylePanel(contentPanel);
        initializeComponents();
    }
    
    protected abstract void initializeComponents();
    protected abstract void onSave();
}