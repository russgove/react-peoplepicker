export enum modes {
    NEW,
    EDIT,
    DISPLAY
}
export class Task {


    public Id: number;
    public Title: string;
    public AssignedToId: number;
    public AssignedTo: string;
    public Priority: string;
    public DueDate: string;
    public Status: string;



}