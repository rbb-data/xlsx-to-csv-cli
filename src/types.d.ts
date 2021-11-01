import { Question as InquirerQuestion } from 'inquirer';

export type Row<T> = Array<T>;
export type Table<T> = Array<Row<T>>;

export type Question = Omit<InquirerQuestion, 'type' | 'name'>;
