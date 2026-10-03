import { JSDOM } from 'jsdom';

const { window } = new JSDOM('');
Object.assign(globalThis, { DOMParser: window.DOMParser, Node: window.Node, Element: window.Element });
