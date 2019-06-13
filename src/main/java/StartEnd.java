/*
StartEnd represents a range of integers between a certain range.
The range should be continuous and there should only be whole
numbers between both numbers.
*/

public class StartEnd {

    private int start;
    private int end;

    public StartEnd() {
        this.start = 0;
        this.end = 0;
    }

    public StartEnd(int start, int end) {
        this.start = start;
        this.end = end;
    }

    public void setStart(int start) {
        this.start = start;
    }

    public void setEnd(int end) {
        this.end = end;
    }

    public int getStart() {
        return start;
    }

    public int getEnd() {
        return end;
    }

    public int getRange() {
        return end - start;
    }
}
