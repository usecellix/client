import React from 'react';
import { createRoot } from 'react-dom/client';
import { AuthGate } from '@/auth';
import '@/auth/auth.css';
import App from './App';

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const container = document.getElementById('root');
    if (container) {
      const root = createRoot(container);
      root.render(
        React.createElement(AuthGate, null, React.createElement(App)),
      );
    }
  }
});

