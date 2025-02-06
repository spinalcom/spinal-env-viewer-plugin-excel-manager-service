export type THeader = {
    key: string | number;
    header: string;
};
export type TSheet = {
    name: string;
    header: THeader[];
    rows: {
        [key: string]: string;
    }[];
};
