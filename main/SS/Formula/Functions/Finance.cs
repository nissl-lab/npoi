using System;
namespace NPOI.SS.Formula.Functions
{


    /**
      * Implementation of the financial functions pmt, fv, ppmt, ipmt.
      * 
      * @author Mike Argyriou micharg@gmail.com
      */
    public class Finance
    {

        /**
         * Emulates Excel/Calc's PMT(interest_rate, number_payments, PV, FV, Type)
         * function, which calculates the payments for a loan or the future value of an investment
         * 
         * @param r
         *            - periodic interest rate represented as a decimal.
         * @param nper
         *            - number of total payments / periods.
         * @param pv
         *            - present value -- borrowed or invested principal.
         * @param fv
         *            - future value of loan or annuity.
         * @param type
         *            - when payment is made: beginning of period is 1; end, 0.
         * @return <code>double</code> representing periodic payment amount.
         */
        // http://arachnoid.com/lutusp/finance.html
        static public double PMT(double r, int nper, double pv, double fv, int type)
        {
            double pmt = -r * (pv * Math.Pow(1 + r, nper) + fv) / ((1 + r * type) * (Math.Pow(1 + r, nper) - 1));
            return pmt;
        }


        /**
         * Overloaded pmt() call omitting type, which defaults to 0.
         * 
         * @see #pmt(double, int, double, double, int)
         */
        static public double PMT(double r, int nper, double pv, double fv)
        {
            return PMT(r, nper, pv, fv, 0);
        }

        /**
         * Overloaded pmt() call omitting fv and type, which both default to 0.
         * 
         * @see #pmt(double, int, double, double, int)
         */
        static public double PMT(double r, int nper, double pv)
        {
            return PMT(r, nper, pv, 0);
        }


        /**
         * Emulates Excel/Calc's IPMT(interest_rate, period, number_payments, PV,
         * FV, Type) function, which calculates the portion of the payment at a
         * given period that is the interest on previous balance.
         * 
         * @param r
         *            - periodic interest rate represented as a decimal.
         * @param per
         *            - period (payment number) to check value at.
         * @param nper
         *            - number of total payments / periods.
         * @param pv
         *            - present value -- borrowed or invested principal.
         * @param fv
         *            - future value of loan or annuity.
         * @param type
         *            - when payment is made: beginning of period is 1; end, 0.
         * @return <code>double</code> representing interest portion of payment.
         * 
         * @see #pmt(double, int, double, double, int)
         * @see #fv(double, int, double, double, int)
         */
        // http://doc.optadata.com/en/dokumentation/application/expression/functions/financial.html
        static public double IPMT(double r, int per, int nper, double pv, double fv, int type)
        {
            double ipmt = FV(r, per - 1, PMT(r, nper, pv, fv, type), pv, type) * r;
            if (type == 1) ipmt /= (1 + r);
            return ipmt;
        }

        static public double IPMT(double r, int per, int nper, double pv, double fv)
        {
            return IPMT(r, per, nper, pv, fv, 0);
        }

        static public double IPMT(double r, int per, int nper, double pv)
        {
            return IPMT(r, per, nper, pv, 0);
        }

        /**
         * Emulates Excel/Calc's PPMT(interest_rate, period, number_payments, PV,
         * FV, Type) function, which calculates the portion of the payment at a
         * given period that will apply to principal.
         * 
         * @param r
         *            - periodic interest rate represented as a decimal.
         * @param per
         *            - period (payment number) to check value at.
         * @param nper
         *            - number of total payments / periods.
         * @param pv
         *            - present value -- borrowed or invested principal.
         * @param fv
         *            - future value of loan or annuity.
         * @param type
         *            - when payment is made: beginning of period is 1; end, 0.
         * @return <code>double</code> representing principal portion of payment.
         * 
         * @see #pmt(double, int, double, double, int)
         * @see #ipmt(double, int, int, double, double, bool)
         */
        static public double PPMT(double r, int per, int nper, double pv, double fv, int type)
        {
            return PMT(r, nper, pv, fv, type) - IPMT(r, per, nper, pv, fv, type);
        }

        static public double PPMT(double r, int per, int nper, double pv, double fv)
        {
            return PMT(r, nper, pv, fv) - IPMT(r, per, nper, pv, fv);
        }

        static public double PPMT(double r, int per, int nper, double pv)
        {
            return PMT(r, nper, pv) - IPMT(r, per, nper, pv);
        }

        /**
         * Emulates Excel/Calc's FV(interest_rate, number_payments, payment, PV,
         * Type) function, which calculates future value or principal at period N.
         * 
         * @param r
         *            - periodic interest rate represented as a decimal.
         * @param nper
         *            - number of total payments / periods.
         * @param pmt
         *            - periodic payment amount.
         * @param pv
         *            - present value -- borrowed or invested principal.
         * @param type
         *            - when payment is made: beginning of period is 1; end, 0.
         * @return <code>double</code> representing future principal value.
         */
        //http://en.wikipedia.org/wiki/Future_value
        static public double FV(double r, int nper, double pmt, double pv, int type)
        {
            double fv = -(pv * Math.Pow(1 + r, nper) + pmt * (1 + r * type) * (Math.Pow(1 + r, nper) - 1) / r);
            return fv;
        }

        /**
         * Overloaded fv() call omitting type, which defaults to 0.
         * 
         * @see #fv(double, int, double, double, int)
         */
        static public double FV(double r, int nper, double c, double pv)
        {
            return FV(r, nper, c, pv, 0);
        }
    }


}