export interface Door {
  page: number;
  section: string;
  drawingCode: string;
  drawingTitle: string;
  group: string;
  name: string;
  color: string;
  type: string;
  doorLeft: string;
  material: string;
  width: string;
  height: string;
  hardware: string;
  singleDouble: string;
  repeat: number;
  comment: string;
}

export interface Room {
  page: number;
  section: string;
  drawingCode: string;
  drawingTitle: string;
  group: string;
  name: string;
  color: string;
  scale: number | undefined;
  ceilingHeight: string;
  area: number;
  exclusionArea: string;
}

export interface Wall {
  page: number;
  section: string;
  drawingCode: string;
  drawingTitle: string;
  group: string;
  name: string;
  color: string;
  scale: number | undefined;
  length: number;
  ceilingHeight: string;
  type: string;
}

export interface Window {
  page: number;
  section: string;
  drawingCode: string;
  drawingTitle: string;
  group: string;
  name: string;
  color: string;
  repeat: number;
  comment: string;
}

export enum ResultAnnotationType {
  WALL = 'Wall',
  ROOM = 'Room',
  WINDOW = 'Window',
  DOOR = 'Door'
}

export interface Point {
  x: number;
  y: number;
}

export type ResultType = Wall | Room | Window | Door;
