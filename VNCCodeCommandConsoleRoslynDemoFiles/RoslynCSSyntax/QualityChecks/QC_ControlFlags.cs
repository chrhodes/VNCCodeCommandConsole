// From Source Code Analysis with Roslyn -

public class QC_ControlFlags
{
    bool cflag = false;

    void update()
    {
        if (!cflag)
            if (thatThing())
                cflag = true;
    }

    bool thatThing()
    {
        return true;
    }

    void thatOtherThing()
    {
        bool flag = false;

        if (flag)
        {
            //Do that
        }
        else
        {
            //Do something else
            flag = false;
        }
    }
}