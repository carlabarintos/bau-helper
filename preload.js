const { contextBridge, ipcRenderer } = require('electron');

// Forward update-available event to renderer
ipcRenderer.on('update-available', (_e, info) => {
  window.dispatchEvent(new CustomEvent('bau-update-available', { detail: info }));
});

contextBridge.exposeInMainWorld('api', {
  az: {
    checkLogin:       ()         => ipcRenderer.invoke('az:checkLogin'),
    login:            ()         => ipcRenderer.invoke('az:login'),
    listSubscriptions:()         => ipcRenderer.invoke('az:listSubscriptions'),
    listWorkspaces:   (subId)    => ipcRenderer.invoke('az:listWorkspaces', subId),
    query:            (params)   => ipcRenderer.invoke('az:query', params),
  },
  ai: {
    summarize: (params) => ipcRenderer.invoke('ai:summarize', params),
  },
  config: {
    load:    ()     => ipcRenderer.invoke('config:load'),
    save:    (data) => ipcRenderer.invoke('config:save', data),
    getPath: ()     => ipcRenderer.invoke('config:getPath'),
  },
  image: {
    save: (params) => ipcRenderer.invoke('image:save', params),
  },
  cache: {
    save:      (rows)  => ipcRenderer.invoke('cache:save', rows),
    clear:     ()      => ipcRenderer.invoke('cache:clear'),
    info:      ()      => ipcRenderer.invoke('cache:info'),
    loadChunk: (index) => ipcRenderer.invoke('cache:loadChunk', index),
  },
});
