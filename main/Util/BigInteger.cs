using System;
using System.Text;
using System.Diagnostics;
using System.Globalization;

namespace NPOI.Util
{
    public class BigInteger : IComparable<BigInteger>
    {
        /**
         * The signum of this BigInteger: -1 for negative, 0 for zero, or
         * 1 for positive.  Note that the BigInteger zero <i>must</i> have
         * a signum of 0.  This is necessary to ensures that there is exactly one
         * representation for each BigInteger value.
         *
         * @serial
         */
        private int _signum;

        /**
         * The magnitude of this BigInteger, in <i>big-endian</i> order: the
         * zeroth element of this array is the most-significant int of the
         * magnitude.  The magnitude must be "minimal" in that the most-significant
         * int ({@code mag[0]}) must be non-zero.  This is necessary to
         * ensure that there is exactly one representation for each BigInteger
         * value.  Note that this implies that the BigInteger zero has a
         * zero-length mag array.
         */
        internal int[] mag;

        // These "redundant fields" are initialized with recognizable nonsense
        // values, and cached the first time they are needed (or never, if they
        // aren't needed).

        /**
        * One plus the bitCount of this BigInteger. Zeros means unitialized.
        *
        * @serial
        * @see #bitCount
        * @deprecated Deprecated since logical value is offset from stored
        * value and correction factor is applied in accessor method.
        */
#if !HIDE_UNREACHABLE_CODE
        [Obsolete]
#endif
        private int bitCount;

        /**
         * One plus the bitLength of this BigInteger. Zeros means unitialized.
         * (either value is acceptable).
         *
         * @serial
         * @see #bitLength()
         * @deprecated Deprecated since logical value is offset from stored
         * value and correction factor is applied in accessor method.
         */
#if !HIDE_UNREACHABLE_CODE
        [Obsolete]
#endif
        private int bitLength;

        // /**
        // * Two plus the lowest set bit of this BigInteger, as returned by
        // * getLowestSetBit().
        // *
        // * @serial
        // * @see #getLowestSetBit
        // * @deprecated Deprecated since logical value is offset from stored
        // * value and correction factor is applied in accessor method.
        // */
       // [Obsolete]
       // never used private int lowestSetBit;

        /**
         * Two plus the index of the lowest-order int in the magnitude of this
         * BigInteger that contains a nonzero int, or -2 (either value is acceptable).
         * The least significant int has int-number 0, the next int in order of
         * increasing significance has int-number 1, and so forth.
         * @deprecated Deprecated since logical value is offset from stored
         * value and correction factor is applied in accessor method.
         */
#if !HIDE_UNREACHABLE_CODE
        [Obsolete]
#endif
        private int firstNonzeroIntNum;

        /**
         * This mask is used to obtain the value of an int as if it were unsigned.
         */
        public const long LONG_MASK = 0xffffffffL;
        public const long INFLATED = long.MinValue;
        public const int Min_RADIX = 2;
        public const int Max_RADIX = 36;
        private static BigInteger[] posConst = new BigInteger[Max_CONSTANT + 1];
        private static BigInteger[] negConst = new BigInteger[Max_CONSTANT + 1];
        private static readonly String[] zeros = new String[64];
        //Constructors
        static BigInteger()
        {
            Init();
        }
        static void Init()
        {
            if (zeros[63] == null)
            {
                for (int i = 1; i <= Max_CONSTANT; i++)
                {
                    int[] magnitude = new int[1];
                    magnitude[0] = i;
                    posConst[i] = new BigInteger(magnitude, 1);
                    negConst[i] = new BigInteger(magnitude, -1);
                }

                zeros[63] = "000000000000000000000000000000000000000000000000000000000000000";
                for (int i = 0; i < 63; i++)
                    zeros[i] = zeros[63].Substring(0, i);
            }
        }
        /**
         * This internal constructor differs from its public cousin
         * with the arguments reversed in two ways: it assumes that its
         * arguments are correct, and it doesn't copy the magnitude array.
         */
        public BigInteger(int[] magnitude, int signum)
        {
            this._signum = (magnitude.Length == 0 ? 0 : signum);
            this.mag = magnitude;
        }
        /**
         * Translates a byte array containing the two's-complement binary
         * representation of a BigInteger into a BigInteger.  The input array is
         * assumed to be in <i>big-endian</i> byte-order: the most significant
         * byte is in the zeroth element.
         *
         * @param  val big-endian two's-complement binary representation of
         *         BigInteger.
         * @throws NumberFormatException {@code val} is zero bytes long.
         */
        public BigInteger(byte[] val)
        {
            if (val.Length == 0)
                throw new ArgumentException("Zero length BigInteger");

            if ((sbyte)val[0] < 0)
            {
                mag = makePositive(val);
                _signum = -1;
            }
            else
            {
                mag = stripLeadingZeroBytes(val);
                _signum = (mag.Length == 0 ? 0 : 1);
            }
        }
        /**
         * This private constructor translates an int array containing the
         * two's-complement binary representation of a BigInteger into a
         * BigInteger. The input array is assumed to be in <i>big-endian</i>
         * int-order: the most significant int is in the zeroth element.
         */
        public BigInteger(int[] val)
        {
            if (val.Length == 0)
                throw new ArgumentException("Zero length BigInteger");

            if (val[0] < 0)
            {
                mag = makePositive(val);
                _signum = -1;
            }
            else
            {
                mag = TrustedStripLeadingZeroInts(val);
                _signum = (mag.Length == 0 ? 0 : 1);
            }
        }
        /**
         * Constructs a BigInteger with the specified value, which may not be zero.
         */
        public BigInteger(long val)
        {
            if (val < 0)
            {
                val = -val;
                _signum = -1;
            }
            else
            {
                _signum = 1;
            }

            int highWord = (int)Operator.UnsignedRightShift(val, 32);
            if (highWord == 0)
            {
                mag = new int[1];
                mag[0] = (int)val;
            }
            else
            {
                mag = new int[2];
                mag[0] = highWord;
                mag[1] = (int)val;
            }
        }

        public BigInteger(string val)
            : this(val, 10)
        {

        }

        public BigInteger(String val, int radix)
        {
            int cursor = 0, numDigits;
            int len = val.Length;

            if (radix < Min_RADIX || radix > Max_RADIX)
                throw new FormatException("Radix out of range");
            if (len == 0)
                throw new FormatException("Zero length BigInteger");

            // Check for at most one leading sign
            int sign = 1;
            int index1 = val.LastIndexOf('-');
            int index2 = val.LastIndexOf('+');
            if ((index1 + index2) <= -1)
            {
                // No leading sign character or at most one leading sign character
                if (index1 == 0 || index2 == 0)
                {
                    cursor = 1;
                    if (len == 1)
                        throw new FormatException("Zero length BigInteger");
                }
                if (index1 == 0)
                    sign = -1;
            }
            else
                throw new FormatException("Illegal embedded sign character");

            // Skip leading zeros and compute number of digits in magnitude
            while (cursor < len &&
                val[cursor] == '0')    //now it can only process 10 radix
                //Character.digit(val.charAt(cursor), radix) == 0) //
                cursor++;
            if (cursor == len)
            {
                _signum = 0;
                mag = ZERO.mag;
                return;
            }

            numDigits = len - cursor;
            _signum = sign;

            // Pre-allocate array of expected size. May be too large but can
            // never be too small. Typically exact.
            int numBits = (int)(Operator.UnsignedRightShift((numDigits * bitsPerDigit[radix]), 10) + 1);
            int numWords = Operator.UnsignedRightShift(numBits + 31, 5);
            int[] magnitude = new int[numWords];

            // Process first (potentially short) digit group
            int firstGroupLen = numDigits % digitsPerInt[radix];
            if (firstGroupLen == 0)
                firstGroupLen = digitsPerInt[radix];
            String group = val.Substring(cursor, cursor += firstGroupLen);
            //magnitude[numWords - 1] = Integer.parseInt(group, radix);
            magnitude[numWords - 1] = int.Parse(group, CultureInfo.InvariantCulture);
            if (magnitude[numWords - 1] < 0)
                throw new FormatException("Illegal digit");

            // Process remaining digit groups
            int superRadix = intRadix[radix];
            int groupVal = 0;
            while (cursor < len)
            {
                group = val.Substring(cursor, cursor += digitsPerInt[radix]);
                //groupVal = Integer.parseInt(group, radix);
                groupVal = int.Parse(group, CultureInfo.InvariantCulture);
                if (groupVal < 0)
                    throw new FormatException("Illegal digit");
                DestructiveMulAdd(magnitude, superRadix, groupVal);
            }
            // Required for cases where the array was overallocated.
            mag = TrustedStripLeadingZeroInts(magnitude);
        }
        /**
         * Returns the input array stripped of any leading zero bytes.
         * Since the source is trusted the copying may be skipped.
         */
        private static int[] TrustedStripLeadingZeroInts(int[] val)
        {
            int vlen = val.Length;
            int keep;

            // Find first nonzero byte
            for (keep = 0; keep < vlen && val[keep] == 0; keep++)
                ;
            return keep == 0 ? val : Arrays.CopyOfRange(val, keep, vlen);
        }

        // Multiply x array times word y in place, and add word z
        private static void DestructiveMulAdd(int[] x, int y, int z)
        {
            // Perform the multiplication word by word
            long ylong = y & LONG_MASK;
            long zlong = z & LONG_MASK;
            int len = x.Length;

            long product = 0;
            long carry = 0;
            for (int i = len - 1; i >= 0; i--)
            {
                product = ylong * (x[i] & LONG_MASK) + carry;
                x[i] = (int)product;
                carry = Operator.UnsignedRightShift(product, 32);
            }

            // Perform the addition
            long sum = (x[len - 1] & LONG_MASK) + zlong;
            x[len - 1] = (int)sum;
            carry = Operator.UnsignedRightShift(sum, 32);
            for (int i = len - 2; i >= 0; i--)
            {
                sum = (x[i] & LONG_MASK) + carry;
                x[i] = (int)sum;
                carry = Operator.UnsignedRightShift(sum, 32);
            }
        }
        /**
         * Returns the String representation of this BigInteger in the
         * given radix.  If the radix is outside the range from {@link
         * Character#Min_RADIX} to {@link Character#Max_RADIX} inclusive,
         * it will default to 10 (as is the case for
         * {@code Integer.toString}).  The digit-to-character mapping
         * provided by {@code Character.forDigit} is used, and a minus
         * sign is prepended if appropriate.  (This representation is
         * compatible with the {@link #BigInteger(String, int) (String,
         * int)} constructor.)
         *
         * @param  radix  radix of the String representation.
         * @return String representation of this BigInteger in the given radix.
         * @see    Integer#toString
         * @see    Character#forDigit
         * @see    #BigInteger(java.lang.String, int)
         */
        public String ToString(int radix)
        {
            if (_signum == 0)
                return "0";
            if (radix < Min_RADIX || radix > Max_RADIX)
                radix = 10;

            //now this method only support 10 radix rendering
            if (radix != 10)
                throw new ArgumentException("Only support 10 radix rendering");

            // Compute upper bound on number of digit groups and allocate space
            int maxNumDigitGroups = (4 * mag.Length + 6) / 7;
            String[] digitGroup = new String[maxNumDigitGroups];

            // Translate number to string, a digit group at a time
            BigInteger tmp = this.Abs();
            int numGroups = 0;
            while (tmp._signum != 0)
            {
                BigInteger d = longRadix[radix];

                MutableBigInteger q = new MutableBigInteger(),
                                  a = new MutableBigInteger(tmp.mag),
                                  b = new MutableBigInteger(d.mag);
                MutableBigInteger r = a.divide(b, q);
                BigInteger q2 = q.toBigInteger(tmp._signum * d._signum);
                BigInteger r2 = r.toBigInteger(tmp._signum * d._signum);

                //digitGroup[numGroups++] = Long.toString(r2.longValue(), radix);
                digitGroup[numGroups++] = r2.LongValue().ToString(CultureInfo.InvariantCulture);
                tmp = q2;
            }

            // Put sign (if any) and first digit group into result buffer
            StringBuilder buf = new StringBuilder(numGroups * digitsPerLong[radix] + 1);
            if (_signum < 0)
                buf.Append('-');
            buf.Append(digitGroup[numGroups - 1]);

            // Append remaining digit groups padded with leading zeros
            for (int i = numGroups - 2; i >= 0; i--)
            {
                // Prepend (any) leading zeros for this digit group
                int numLeadingZeros = digitsPerLong[radix] - digitGroup[i].Length;
                if (numLeadingZeros != 0)
                    buf.Append(zeros[numLeadingZeros]);
                buf.Append(digitGroup[i]);
            }
            return buf.ToString();
        }
        /**
         * The BigInteger constant zero.
         *
         * @since   1.2
         */
        public static readonly BigInteger ZERO = new BigInteger(new int[0], 0);

        /**
         * The BigInteger constant one.
         *
         * @since   1.2
         */
        public static readonly BigInteger One = ValueOf(1);

        /**
         * The BigInteger constant two.  (Not exported.)
         */
        private static readonly BigInteger Two = ValueOf(2);

        /**
         * The BigInteger constant ten.
         *
         * @since   1.5
         */
        public static readonly BigInteger TEN = ValueOf(10);

        /**
         * Returns a BigInteger whose value is equal to that of the
         * specified {@code long}.  This "static factory method" is
         * provided in preference to a ({@code long}) constructor
         * because it allows for reuse of frequently used BigIntegers.
         *
         * @param  val value of the BigInteger to return.
         * @return a BigInteger with the specified value.
         */
        public static BigInteger ValueOf(long val)
        {
            Init();
           // If -Max_CONSTANT < val < Max_CONSTANT, return stashed constant
            if (val == 0)
                return ZERO;
            if (val > 0 && val <= Max_CONSTANT)
                return posConst[(int)val];
            else if (val < 0 && val >= -Max_CONSTANT)
                return negConst[(int)-val];

            return new BigInteger(val);
        }
        private const int Max_CONSTANT = 16;

        /**
         * Returns a BigInteger with the given two's complement representation.
         * Assumes that the input array will not be modified (the returned
         * BigInteger will reference the input array if feasible).
         */
        private static BigInteger ValueOf(int[] val)
        {
            return (val[0] > 0 ? new BigInteger(val, 1) : new BigInteger(val));
        }
        /**
         * Package private method to return bit length for an integer.
         */
        public static int BitLengthForInt(int n)
        {
            return 32 - NumberOfLeadingZeros(n);
        }
        // Miscellaneous Bit Operations

        /*
         * Returns the number of bits in the minimal two's-complement
         * representation of this BigInteger, <i>excluding</i> a sign bit.
         * For positive BigIntegers, this is equivalent to the number of bits in
         * the ordinary binary representation.  (Computes
         * <c>(ceil(log2(this <&lt; 0 ? -this : this+1)))</c>.)
         *
         * @return number of bits in the minimal two's-complement
         *         representation of this BigInteger, <i>excluding</i> a sign bit.
         */
        public int BitLength()
        {
            int n = bitLength - 1;
            if (n == -1)
            { // bitLength not initialized yet
                int[] m = mag;
                int len = m.Length;
                if (len == 0)
                {
                    n = 0; // offset by one to initialize
                }
                else
                {
                    // Calculate the bit length of the magnitude
                    int magBitLength = ((len - 1) << 5) + BitLengthForInt(mag[0]);
                    if (_signum < 0)
                    {
                        // Check if magnitude is a power of two
                        bool pow2 = (BitCountForInt(mag[0]) == 1);
                        for (int i = 1; i < len && pow2; i++)
                            pow2 = (mag[i] == 0);

                        n = (pow2 ? magBitLength - 1 : magBitLength);
                    }
                    else
                    {
                        n = magBitLength;
                    }
                }
                bitLength = n + 1;
            }
            return n;
        }
        /**
         * Returns the number of bits in the two's complement representation
         * of this BigInteger that differ from its sign bit.  This method is
         * useful when implementing bit-vector style sets atop BigIntegers.
         *
         * @return number of bits in the two's complement representation
         *         of this BigInteger that differ from its sign bit.
         */
        public int BitCount()
        {
            //@SuppressWarnings("deprecation") 
            int bc = bitCount - 1;
            if (bc == -1)
            {  // bitCount not initialized yet
                bc = 0;      // offset by one to initialize
                // Count the bits in the magnitude
                for (int i = 0; i < mag.Length; i++)
                    bc += BitCountForInt(mag[i]);
                if (_signum < 0)
                {
                    // Count the trailing zeros in the magnitude
                    int magTrailingZeroCount = 0, j;
                    for (j = mag.Length - 1; mag[j] == 0; j--)
                        magTrailingZeroCount += 32;
                    magTrailingZeroCount += NumberOfTrailingZeros(mag[j]);
                    bc += magTrailingZeroCount - 1;
                }
                bitCount = bc + 1;
            }
            return bc;
        }

        /**
         * Returns a BigInteger whose value is the absolute value of this
         * BigInteger.
         *
         * @return {@code abs(this)}
         */
        public BigInteger Abs()
        {
            return (_signum >= 0 ? this : this.Negate());
        }

        /**
         * Returns a BigInteger whose value is {@code (-this)}.
         *
         * @return {@code -this}
         */
        public BigInteger Negate()
        {
            return new BigInteger(this.mag, -this._signum);
        }
        /**
         * Returns a BigInteger whose value is <c>(this<sup>exponent</sup>)</c>.
         * Note that {@code exponent} is an integer rather than a BigInteger.
         *
         * @param  exponent exponent to which this BigInteger is to be raised.
         * @return <c>this<sup>exponent</sup></c>
         * @throws ArithmeticException {@code exponent} is negative.  (This would
         *         cause the operation to yield a non-integer value.)
         */
        public BigInteger Pow(int exponent)
        {
            if (exponent < 0)
                throw new ArithmeticException("Negative exponent");
            if (_signum == 0)
                return (exponent == 0 ? One : this);

            // Perform exponentiation using repeated squaring trick
            int newSign = (_signum < 0 && (exponent & 1) == 1 ? -1 : 1);
            int[] baseToPow2 = this.mag;
            int[] result = { 1 };

            while (exponent != 0)
            {
                if ((exponent & 1) == 1)
                {
                    result = MultiplyToLen(result, result.Length,
                                           baseToPow2, baseToPow2.Length, null);
                    result = TrustedStripLeadingZeroInts(result);
                }
                exponent = Operator.UnsignedRightShift(exponent, 1);
                if (exponent != 0)
                {
                    baseToPow2 = squareToLen(baseToPow2, baseToPow2.Length, null);
                    baseToPow2 = TrustedStripLeadingZeroInts(baseToPow2);
                }
            }
            return new BigInteger(result, newSign);
        }
        /**
         * Multiplies int arrays x and y to the specified lengths and places
         * the result into z. There will be no leading zeros in the resultant array.
         */
        private int[] MultiplyToLen(int[] x, int xlen, int[] y, int ylen, int[] z)
        {
            int xstart = xlen - 1;
            int ystart = ylen - 1;

            if (z == null || z.Length < (xlen + ylen))
                z = new int[xlen + ylen];

            long carry = 0;
            for (int j = ystart, k = ystart + 1 + xstart; j >= 0; j--, k--)
            {
                long product = (y[j] & LONG_MASK) *
                               (x[xstart] & LONG_MASK) + carry;
                z[k] = (int)product;
                carry = Operator.UnsignedRightShift(product, 32);
            }
            z[xstart] = (int)carry;

            for (int i = xstart - 1; i >= 0; i--)
            {
                carry = 0;
                for (int j = ystart, k = ystart + 1 + i; j >= 0; j--, k--)
                {
                    long product = (y[j] & LONG_MASK) *
                                   (x[i] & LONG_MASK) +
                                   (z[k] & LONG_MASK) + carry;
                    z[k] = (int)product;
                    carry = Operator.UnsignedRightShift(product, 32);
                }
                z[i] = (int)carry;
            }
            return z;
        }
        /**
         * Multiply an array by one word k and add to result, return the carry
         */
        static int mulAdd(int[] output, int[] input, int offset, int len, int k)
        {
            long kLong = k & LONG_MASK;
            long carry = 0;

            offset = output.Length - offset - 1;
            for (int j = len - 1; j >= 0; j--)
            {
                long product = (input[j] & LONG_MASK) * kLong +
                               (output[offset] & LONG_MASK) + carry;
                output[offset--] = (int)product;
                carry = Operator.UnsignedRightShift(product, 32);
            }
            return (int)carry;
        }
        /**
         * Squares the contents of the int array x. The result is placed into the
         * int array z.  The contents of x are not changed.
         */
        private static int[] squareToLen(int[] x, int len, int[] z)
        {
            /*
             * The algorithm used here is adapted from Colin Plumb's C library.
             * Technique: Consider the partial products in the multiplication
             * of "abcde" by itself:
             *
             *               a  b  c  d  e
             *            *  a  b  c  d  e
             *          ==================
             *              ae be ce de ee
             *           ad bd cd dd de
             *        ac bc cc cd ce
             *     ab bb bc bd be
             *  aa ab ac ad ae
             *
             * Note that everything above the main diagonal:
             *              ae be ce de = (abcd) * e
             *           ad bd cd       = (abc) * d
             *        ac bc             = (ab) * c
             *     ab                   = (a) * b
             *
             * is a copy of everything below the main diagonal:
             *                       de
             *                 cd ce
             *           bc bd be
             *     ab ac ad ae
             *
             * Thus, the sum is 2 * (off the diagonal) + diagonal.
             *
             * This is accumulated beginning with the diagonal (which
             * consist of the squares of the digits of the input), which is then
             * divided by two, the off-diagonal added, and multiplied by two
             * again.  The low bit is simply a copy of the low bit of the
             * input, so it doesn't need special care.
             */
            int zlen = len << 1;
            if (z == null || z.Length < zlen)
                z = new int[zlen];

            // Store the squares, right shifted one bit (i.e., divided by 2)
            int lastProductLowWord = 0;
            for (int j = 0, i = 0; j < len; j++)
            {
                long piece = (x[j] & LONG_MASK);
                long product = piece * piece;
                z[i++] = (lastProductLowWord << 31) | (int)Operator.UnsignedRightShift(product, 33);
                z[i++] = (int)Operator.UnsignedRightShift(product, 1);
                lastProductLowWord = (int)product;
            }

            // Add in off-diagonal sums
            for (int i = len, offset = 1; i > 0; i--, offset += 2)
            {
                int t = x[i - 1];
                t = mulAdd(z, x, offset, i - 1, t);
                addOne(z, offset - 1, i, t);
            }

            // Shift back up and set low bit
            PrimitiveLeftShift(z, zlen, 1);
            z[zlen - 1] |= x[len - 1] & 1;

            return z;
        }
        // shifts a up to len left n bits assumes no leading zeros, 0<=n<32
        public static void PrimitiveLeftShift(int[] a, int len, int n)
        {
            if (len == 0 || n == 0)
                return;

            int n2 = 32 - n;
            for (int i = 0, c = a[i], m = i + len - 1; i < m; i++)
            {
                int b = c;
                c = a[i + 1];
                a[i] = (b << n) | Operator.UnsignedRightShift(c, n2);
            }
            a[len - 1] <<= n;
        }
        /**
         * Add one word to the number a mlen words into a. Return the resulting
         * carry.
         */
        static int addOne(int[] a, int offset, int mlen, int carry)
        {
            offset = a.Length - 1 - mlen - offset;
            long t = (a[offset] & LONG_MASK) + (carry & LONG_MASK);

            a[offset] = (int)t;
            if ((t >> 32) == 0)
                return 0;
            while (--mlen >= 0)
            {
                if (--offset < 0)
                { // Carry out of number
                    return 1;
                }
                else
                {
                    a[offset]++;
                    if (a[offset] != 0)
                        return 0;
                }
            }
            return 1;
        }
        /**
         * Returns the signum function of this BigInteger.
         *
         * @return -1, 0 or 1 as the value of this BigInteger is negative, zero or
         *         positive.
         */
        public int Signum()
        {
            return this._signum;
        }
        /**
         * Returns a byte array containing the two's-complement
         * representation of this BigInteger.  The byte array will be in
         * <i>big-endian</i> byte-order: the most significant byte is in
         * the zeroth element.  The array will contain the minimum number
         * of bytes required to represent this BigInteger, including at
         * least one sign bit, which is {@code (ceil((this.bitLength() +
         * 1)/8))}.  (This representation is compatible with the
         * {@link #BigInteger(byte[]) (byte[])} constructor.)
         *
         * @return a byte array containing the two's-complement representation of
         *         this BigInteger.
         * @see    #BigInteger(byte[])
         */
        public byte[] ToByteArray()
        {
            int byteLen = (BitLength() / 8 + 1);
            byte[] byteArray = new byte[byteLen];

            for (int i = byteLen - 1, bytesCopied = 4, nextInt = 0, intIndex = 0; i >= 0; i--)
            {
                if (bytesCopied == 4)
                {
                    nextInt = GetInt(intIndex++);
                    bytesCopied = 1;
                }
                else
                {
                    nextInt = Operator.UnsignedRightShift(nextInt, 8);
                    bytesCopied++;
                }
                byteArray[i] = (byte)nextInt;
            }
            return byteArray;
        }
        /**
         * Returns the length of the two's complement representation in ints,
         * including space for at least one sign bit.
         */
        private int intLength()
        {
            return Operator.UnsignedRightShift(BitLength(), 5) + 1;
        }

        /* Returns sign bit */
        private int signBit()
        {
            return _signum < 0 ? 1 : 0;
        }

        /* Returns an int of sign bits */
        private int signInt()
        {
            return _signum < 0 ? -1 : 0;
        }
        /**
         * Returns the specified int of the little-endian two's complement
         * representation (int 0 is the least significant).  The int number can
         * be arbitrarily high (values are logically preceded by infinitely many
         * sign ints).
         */
        private int GetInt(int n)
        {
            if (n < 0)
                return 0;
            if (n >= mag.Length)
                return signInt();

            int magInt = mag[mag.Length - n - 1];

            return (_signum >= 0 ? magInt :
                    (n <= FirstNonzeroIntNum() ? -magInt : ~magInt));
        }
        /**
         * Returns the index of the int that contains the first nonzero int in the
         * little-endian binary representation of the magnitude (int 0 is the
         * least significant). If the magnitude is zero, return value is undefined.
         */
        private int FirstNonzeroIntNum()
        {
            int fn = firstNonzeroIntNum - 2;
            if (fn == -2)
            { // firstNonzeroIntNum not initialized yet
                fn = 0;

                // Search for the first nonzero int
                int i;
                int mlen = mag.Length;
                for (i = mlen - 1; i >= 0 && mag[i] == 0; i--)
                    ;
                fn = mlen - i - 1;
                firstNonzeroIntNum = fn + 2; // offset by two to initialize
            }
            return fn;
        }
        /**
         * Returns a copy of the input array stripped of any leading zero bytes.
         */
        private static int[] stripLeadingZeroBytes(byte[] a)
        {
            int byteLength = a.Length;
            int keep;

            // Find first nonzero byte
            for (keep = 0; keep < byteLength && a[keep] == 0; keep++)
                ;

            // Allocate new array and copy relevant part of input array
            int intLength = Operator.UnsignedRightShift((byteLength - keep) + 3, 2);
            int[] result = new int[intLength];
            int b = byteLength - 1;
            for (int i = intLength - 1; i >= 0; i--)
            {
                result[i] = a[b--] & 0xff;
                int bytesRemaining = b - keep + 1;
                int bytesToTransfer = Math.Min(3, bytesRemaining);
                for (int j = 8; j <= (bytesToTransfer << 3); j += 8)
                    result[i] |= ((a[b--] & 0xff) << j);
            }
            return result;
        }
        /**
         * Takes an array a representing a negative 2's-complement number and
         * returns the minimal (no leading zero bytes) unsigned whose value is -a.
         */
        private static int[] makePositive(byte[] a)
        {
            int keep, k;
            int byteLength = a.Length;

            // Find first non-sign (0xff) byte of input
            for (keep = 0; (keep < byteLength) && ((sbyte)a[keep] == (sbyte)-1); keep++)
                ;


            /* Allocate output array.  If all non-sign bytes are 0x00, we must
             * allocate space for one extra output byte. */
            for (k = keep; k < byteLength && a[k] == 0; k++)
                ;

            int extraByte = (k == byteLength) ? 1 : 0;
            int intLength = ((byteLength - keep + extraByte) + 3) / 4;
            int[] result = new int[intLength];

            /* Copy one's complement of input into output, leaving extra
             * byte (if it exists) == 0x00 */
            int b = byteLength - 1;
            for (int i = intLength - 1; i >= 0; i--)
            {
                result[i] = a[b--] & 0xff;
                int numBytesToTransfer = Math.Min(3, b - keep + 1);
                if (numBytesToTransfer < 0)
                    numBytesToTransfer = 0;
                for (int j = 8; j <= 8 * numBytesToTransfer; j += 8)
                    result[i] |= ((a[b--] & 0xff) << j);

                // Mask indicates which bits must be complemented
                int mask = Operator.UnsignedRightShift(-1, (8 * (3 - numBytesToTransfer)));
                result[i] = ~result[i] & mask;
            }

            // Add one to one's complement to generate two's complement
            for (int i = result.Length - 1; i >= 0; i--)
            {
                result[i] = (int)((result[i] & LONG_MASK) + 1);
                if (result[i] != 0)
                    break;
            }

            return result;
        }
        /*
         * The following two arrays are used for fast String conversions.  Both
         * are indexed by radix.  The first is the number of digits of the given
         * radix that can fit in a Java long without "going negative", i.e., the
         * highest integer n such that radix**n < 2**63.  The second is the
         * "long radix" that tears each number into "long digits", each of which
         * consists of the number of digits in the corresponding element in
         * digitsPerLong (longRadix[i] = i**digitPerLong[i]).  Both arrays have
         * nonsense values in their 0 and 1 elements, as radixes 0 and 1 are not
         * used.
         */
        private static readonly int[] digitsPerLong = {0, 0,
        62, 39, 31, 27, 24, 22, 20, 19, 18, 18, 17, 17, 16, 16, 15, 15, 15, 14,
        14, 14, 14, 13, 13, 13, 13, 13, 13, 12, 12, 12, 12, 12, 12, 12, 12};

        private static readonly BigInteger[] longRadix = {null, null,
        ValueOf(0x4000000000000000L), ValueOf(0x383d9170b85ff80bL),
        ValueOf(0x4000000000000000L), ValueOf(0x6765c793fa10079dL),
        ValueOf(0x41c21cb8e1000000L), ValueOf(0x3642798750226111L),
        ValueOf(0x1000000000000000L), ValueOf(0x12bf307ae81ffd59L),
        ValueOf( 0xde0b6b3a7640000L), ValueOf(0x4d28cb56c33fa539L),
        ValueOf(0x1eca170c00000000L), ValueOf(0x780c7372621bd74dL),
        ValueOf(0x1e39a5057d810000L), ValueOf(0x5b27ac993df97701L),
        ValueOf(0x1000000000000000L), ValueOf(0x27b95e997e21d9f1L),
        ValueOf(0x5da0e1e53c5c8000L), ValueOf( 0xb16a458ef403f19L),
        ValueOf(0x16bcc41e90000000L), ValueOf(0x2d04b7fdd9c0ef49L),
        ValueOf(0x5658597bcaa24000L), ValueOf( 0x6feb266931a75b7L),
        ValueOf( 0xc29e98000000000L), ValueOf(0x14adf4b7320334b9L),
        ValueOf(0x226ed36478bfa000L), ValueOf(0x383d9170b85ff80bL),
        ValueOf(0x5a3c23e39c000000L), ValueOf( 0x4e900abb53e6b71L),
        ValueOf( 0x7600ec618141000L), ValueOf( 0xaee5720ee830681L),
        ValueOf(0x1000000000000000L), ValueOf(0x172588ad4f5f0981L),
        ValueOf(0x211e44f7d02c1000L), ValueOf(0x2ee56725f06e5c71L),
        ValueOf(0x41c21cb8e1000000L)};

        // bitsPerDigit in the given radix times 1024
        // Rounded up to avoid underallocation.
        private static readonly long[] bitsPerDigit = { 0, 0,
        1024, 1624, 2048, 2378, 2648, 2875, 3072, 3247, 3402, 3543, 3672,
        3790, 3899, 4001, 4096, 4186, 4271, 4350, 4426, 4498, 4567, 4633,
        4696, 4756, 4814, 4870, 4923, 4975, 5025, 5074, 5120, 5166, 5210,
                                           5253, 5295};
        /*
         * These two arrays are the integer analogue of above.
         */
        private static readonly int[] digitsPerInt = {0, 0, 30, 19, 15, 13, 11,
        11, 10, 9, 9, 8, 8, 8, 8, 7, 7, 7, 7, 7, 7, 7, 6, 6, 6, 6,
        6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 5};

        private static readonly int[] intRadix = {0, 0,
        0x40000000, 0x4546b3db, 0x40000000, 0x48c27395, 0x159fd800,
        0x75db9c97, 0x40000000, 0x17179149, 0x3b9aca00, 0xcc6db61,
        0x19a10000, 0x309f1021, 0x57f6c100, 0xa2f1b6f,  0x10000000,
        0x18754571, 0x247dbc80, 0x3547667b, 0x4c4b4000, 0x6b5a6e1d,
        0x6c20a40,  0x8d2d931,  0xb640000,  0xe8d4a51,  0x1269ae40,
        0x17179149, 0x1cb91000, 0x23744899, 0x2b73a840, 0x34e63b41,
        0x40000000, 0x4cfa3cc1, 0x5c13d840, 0x6d91b519, 0x39aa400
    };
        /**
         * Takes an array a representing a negative 2's-complement number and
         * returns the minimal (no leading zero ints) unsigned whose value is -a.
         */
        private static int[] makePositive(int[] a)
        {
            int keep, j;

            // Find first non-sign (0xffffffff) int of input
            for (keep = 0; keep < a.Length && a[keep] == -1; keep++)
                ;

            /* Allocate output array.  If all non-sign ints are 0x00, we must
             * allocate space for one extra output int. */
            for (j = keep; j < a.Length && a[j] == 0; j++)
                ;
            int extraInt = (j == a.Length ? 1 : 0);
            int[] result = new int[a.Length - keep + extraInt];

            /* Copy one's complement of input into output, leaving extra
             * int (if it exists) == 0x00 */
            for (int i = keep; i < a.Length; i++)
                result[i - keep + extraInt] = ~a[i];

            // Add one to one's complement to generate two's complement
            for (int i = result.Length - 1; ++result[i] == 0; i--)
                ;

            return result;
        }
        #region Integer Method
        /**
         * Returns the number of zero bits preceding the highest-order
         * ("leftmost") one-bit in the two's complement binary representation
         * of the specified {@code int} value.  Returns 32 if the
         * specified value has no one-bits in its two's complement representation,
         * in other words if it is equal to zero.
         *
         * Note that this method is closely related to the logarithm base 2.
         * For all positive {@code int} values x:
         * <ul>
         * <li>floor(log<sub>2</sub>(x)) = {@code 31 - numberOfLeadingZeros(x)}</li>
         * <li>ceil(log<sub>2</sub>(x)) = {@code 32 - numberOfLeadingZeros(x - 1)}</li>
         * </ul>
         *
         * @return the number of zero bits preceding the highest-order
         *     ("leftmost") one-bit in the two's complement binary representation
         *     of the specified {@code int} value, or 32 if the value
         *     is equal to zero.
         * @since 1.5
         */
        public static int NumberOfLeadingZeros(int i)
        {
            // HD, Figure 5-6
            if (i == 0)
                return 32;
            int n = 1;
            if (Operator.UnsignedRightShift(i, 16) == 0) { n += 16; i <<= 16; }
            if (Operator.UnsignedRightShift(i, 24) == 0) { n += 8; i <<= 8; }
            if (Operator.UnsignedRightShift(i, 28) == 0) { n += 4; i <<= 4; }
            if (Operator.UnsignedRightShift(i, 30) == 0) { n += 2; i <<= 2; }
            n -= Operator.UnsignedRightShift(i, 31);
            return n;
        }
        /**
         * Returns the number of zero bits following the lowest-order ("rightmost")
         * one-bit in the two's complement binary representation of the specified
         * {@code int} value.  Returns 32 if the specified value has no
         * one-bits in its two's complement representation, in other words if it is
         * equal to zero.
         *
         * @return the number of zero bits following the lowest-order ("rightmost")
         *     one-bit in the two's complement binary representation of the
         *     specified {@code int} value, or 32 if the value is equal
         *     to zero.
         * @since 1.5
         */
        public static int NumberOfTrailingZeros(int i)
        {
            // HD, Figure 5-14
            int y;
            if (i == 0) return 32;
            int n = 31;
            y = i << 16; if (y != 0) { n = n - 16; i = y; }
            y = i << 8; if (y != 0) { n = n - 8; i = y; }
            y = i << 4; if (y != 0) { n = n - 4; i = y; }
            y = i << 2; if (y != 0) { n = n - 2; i = y; }
            return n - Operator.UnsignedRightShift((i << 1), 31);
        }
        /**
         * Returns the number of one-bits in the two's complement binary
         * representation of the specified {@code int} value.  This function is
         * sometimes referred to as the <i>population count</i>.
         *
         * @return the number of one-bits in the two's complement binary
         *     representation of the specified {@code int} value.
         * @since 1.5
         */
        public static int BitCountForInt(int i)
        {
            // HD, Figure 5-2
            uint x = (uint)i;
            x = x - ((x >> 1) & 0x55555555);
            x = (x & 0x33333333) + ((x >> 2) & 0x33333333);
            x = (x + (x >> 4)) & 0x0f0f0f0f;
            x = x + (x >> 8);
            x = x + (x >> 16);
            return (int)(x & 0x3f);
        }
        #endregion
        #region IComparable<BigInteger> 成员

        public int CompareTo(BigInteger val)
        {
            if (_signum == val._signum)
            {
                switch (_signum)
                {
                    case 1:
                        return compareMagnitude(val);
                    case -1:
                        return val.compareMagnitude(this);
                    default:
                        return 0;
                }
            }
            return _signum > val._signum ? 1 : -1;
        }
        /**
     * Compares the magnitude array of this BigInteger with the specified
     * BigInteger's. This is the version of compareTo ignoring sign.
     *
     * @param val BigInteger whose magnitude array to be compared.
     * @return -1, 0 or 1 as this magnitude array is less than, equal to or
     *         greater than the magnitude aray for the specified BigInteger's.
     */
        int compareMagnitude(BigInteger val)
        {
            int[] m1 = mag;
            int len1 = m1.Length;
            int[] m2 = val.mag;
            int len2 = m2.Length;
            if (len1 < len2)
                return -1;
            if (len1 > len2)
                return 1;
            for (int i = 0; i < len1; i++)
            {
                int a = m1[i];
                int b = m2[i];
                if (a != b)
                    return ((a & LONG_MASK) < (b & LONG_MASK)) ? -1 : 1;
            }
            return 0;
        }
        #endregion

        /**
         * Compares this BigInteger with the specified Object for equality.
         *
         * @param  x Object to which this BigInteger is to be compared.
         * @return {@code true} if and only if the specified Object is a
         *         BigInteger whose value is numerically equal to this BigInteger.
         */
        public override bool Equals(object x)
        {
            // This test is just an optimization, which may or may not help
            //if (x == this) - avoid CS0252 by making the reference comparison explicit:
            if (Object.ReferenceEquals(x, this))
                return true;

            if (!(x is BigInteger) || (null == x))
                return false;

            BigInteger xInt = (BigInteger)x;
            if (xInt._signum != _signum)
                return false;

            int[] m = mag;
            int len = m.Length;
            int[] xm = xInt.mag;
            if (len != xm.Length)
                return false;

            for (int i = 0; i < len; i++)
                if (xm[i] != m[i])
                    return false;

            return true;
        }
        /**
         * Returns the minimum of this BigInteger and {@code val}.
         *
         * @param  val value with which the minimum is to be computed.
         * @return the BigInteger whose value is the lesser of this BigInteger and
         *         {@code val}.  If they are equal, either may be returned.
         */
        public BigInteger Min(BigInteger val)
        {
            return (CompareTo(val) < 0 ? this : val);
        }

        /**
         * Returns the maximum of this BigInteger and {@code val}.
         *
         * @param  val value with which the maximum is to be computed.
         * @return the BigInteger whose value is the greater of this and
         *         {@code val}.  If they are equal, either may be returned.
         */
        public BigInteger Max(BigInteger val)
        {
            return (CompareTo(val) > 0 ? this : val);
        }


        // Hash Function

        /**
         * Returns the hash code for this BigInteger.
         *
         * @return hash code for this BigInteger.
         */
        public override int GetHashCode()
        {
            int hashCode = 0;

            for (int i = 0; i < mag.Length; i++)
                hashCode = (int)(31 * hashCode + (mag[i] & LONG_MASK));

            return hashCode * _signum;
        }
        /**
         * Converts this BigInteger to an {@code int}.  This
         * conversion is analogous to a
         * <i>narrowing primitive conversion</i> from {@code long} to
         * {@code int} as defined in section 5.1.3 of
         * <cite>The Java(TM) Language Specification</cite>:
         * if this BigInteger is too big to fit in an
         * {@code int}, only the low-order 32 bits are returned.
         * Note that this conversion can lose information about the
         * overall magnitude of the BigInteger value as well as return a
         * result with the opposite sign.
         *
         * @return this BigInteger converted to an {@code int}.
         */
        public int IntValue()
        {
            int result = 0;
            result = GetInt(0);
            return result;
        }
        public BigInteger ShiftLeft(int n)
        {
            if (_signum == 0)
                return ZERO;
            if (n == 0)
                return this;
            if (n < 0)
            {
                if (n == int.MinValue)
                {
                    throw new ArithmeticException("Shift distance of Integer.Min_VALUE not supported.");
                }
                else
                {
                    return ShiftRight(-n);
                }
            }

            int nInts = Operator.UnsignedRightShift(n, 5);
            int nBits = n & 0x1f;
            int magLen = mag.Length;
            int[] newMag = null;

            if (nBits == 0)
            {
                newMag = new int[magLen + nInts];
                for (int i = 0; i < magLen; i++)
                    newMag[i] = mag[i];
            }
            else
            {
                int i = 0;
                int nBits2 = 32 - nBits;
                int highBits = Operator.UnsignedRightShift(mag[0], nBits2);
                if (highBits != 0)
                {
                    newMag = new int[magLen + nInts + 1];
                    newMag[i++] = highBits;
                }
                else
                {
                    newMag = new int[magLen + nInts];
                }
                int j = 0;
                while (j < magLen - 1)
                    newMag[i++] = mag[j++] << nBits | Operator.UnsignedRightShift(mag[j], nBits2);
                newMag[i] = mag[j] << nBits;
            }

            return new BigInteger(newMag, _signum);
        }
        /**
         * Converts this BigInteger to a {@code long}.  This
         * conversion is analogous to a
         * <i>narrowing primitive conversion</i> from {@code long} to
         * {@code int} as defined in section 5.1.3 of
         * <cite>The Java(TM) Language Specification</cite>:
         * if this BigInteger is too big to fit in a
         * {@code long}, only the low-order 64 bits are returned.
         * Note that this conversion can lose information about the
         * overall magnitude of the BigInteger value as well as return a
         * result with the opposite sign.
         *
         * @return this BigInteger converted to a {@code long}.
         */
        public long LongValue()
        {
            long result = 0;

            for (int i = 1; i >= 0; i--)
                result = (result << 32) + (GetInt(i) & LONG_MASK);
            return result;
        }
        /**
         * Returns a BigInteger whose value is {@code (this >> n)}.  Sign
         * extension is performed.  The shift distance, {@code n}, may be
         * negative, in which case this method performs a left shift.
         * (Computes <c>floor(this / 2<sup>n</sup>)</c>.)
         *
         * @param  n shift distance, in bits.
         * @return {@code this >> n}
         * @throws ArithmeticException if the shift distance is {@code
         *         Integer.Min_VALUE}.
         * @see #shiftLeft
         */
        public BigInteger ShiftRight(int n)
        {
            if (n == 0)
                return this;
            if (n < 0)
            {
                if (n == int.MinValue)
                {
                    throw new ArithmeticException("Shift distance of Integer.Min_VALUE not supported.");
                }
                else
                {
                    return ShiftLeft(-n);
                }
            }

            int nInts = Operator.UnsignedRightShift(n, 5);
            int nBits = n & 0x1f;
            int magLen = mag.Length;
            int[] newMag = null;

            // Special case: entire contents shifted off the end
            if (nInts >= magLen)
                return (_signum >= 0 ? ZERO : negConst[1]);

            if (nBits == 0)
            {
                int newMagLen = magLen - nInts;
                newMag = new int[newMagLen];
                for (int i = 0; i < newMagLen; i++)
                    newMag[i] = mag[i];
            }
            else
            {
                int i = 0;
                int highBits = Operator.UnsignedRightShift(mag[0], nBits);
                if (highBits != 0)
                {
                    newMag = new int[magLen - nInts];
                    newMag[i++] = highBits;
                }
                else
                {
                    newMag = new int[magLen - nInts - 1];
                }

                int nBits2 = 32 - nBits;
                int j = 0;
                while (j < magLen - nInts - 1)
                    newMag[i++] = (mag[j++] << nBits2) | Operator.UnsignedRightShift(mag[j], nBits);
            }

            if (_signum < 0)
            {
                // Find out whether any one-bits were shifted off the end.
                bool onesLost = false;
                for (int i = magLen - 1, j = magLen - nInts; i >= j && !onesLost; i--)
                    onesLost = (mag[i] != 0);
                if (!onesLost && nBits != 0)
                    onesLost = (mag[magLen - nInts - 1] << (32 - nBits) != 0);

                if (onesLost)
                    newMag = Increment(newMag);
            }

            return new BigInteger(newMag, _signum);
        }
        int[] Increment(int[] val)
        {
            int lastSum = 0;
            for (int i = val.Length - 1; i >= 0 && lastSum == 0; i--)
                lastSum = (val[i] += 1);
            if (lastSum == 0)
            {
                val = new int[val.Length + 1];
                val[0] = 1;
            }
            return val;
        }
        public BigInteger and(BigInteger val)
        {
            int[] result = new int[Math.Max(intLength(), val.intLength())];
            for (int i = 0; i < result.Length; i++)
                result[i] = (GetInt(result.Length - i - 1)
                             & val.GetInt(result.Length - i - 1));

            return ValueOf(result);
        }
        /**
         * Returns a BigInteger whose value is {@code (~this)}.  (This method
         * returns a negative value if and only if this BigInteger is
         * non-negative.)
         *
         * @return {@code ~this}
         */
        public BigInteger Not()
        {
            int[] result = new int[intLength()];
            for (int i = 0; i < result.Length; i++)
                result[i] = ~GetInt(result.Length - i - 1);

            return ValueOf(result);
        }
        /**
         * Returns a BigInteger whose value is {@code (this | val)}.  (This method
         * returns a negative BigInteger if and only if either this or val is
         * negative.)
         *
         * @param val value to be OR'ed with this BigInteger.
         * @return {@code this | val}
         */
        public BigInteger Or(BigInteger val)
        {
            int[] result = new int[Math.Max(intLength(), val.intLength())];
            for (int i = 0; i < result.Length; i++)
                result[i] = (GetInt(result.Length - i - 1)
                             | val.GetInt(result.Length - i - 1));

            return ValueOf(result);
        }
        /**
         * Package private methods used by BigDecimal code to multiply a BigInteger
         * with a long. Assumes v is not equal to INFLATED.
         */
        BigInteger Multiply(long v)
        {
            if (v == 0 || _signum == 0)
                return ZERO;
            if (v == INFLATED)
                return Multiply(BigInteger.ValueOf(v));
            int rsign = (v > 0 ? _signum : -_signum);
            if (v < 0)
                v = -v;
            long dh = Operator.UnsignedRightShift(v, 32);      // higher order bits
            long dl = v & LONG_MASK; // lower order bits

            int xlen = mag.Length;
            int[] value = mag;
            int[] rmag = (dh == 0L) ? (new int[xlen + 1]) : (new int[xlen + 2]);
            long carry = 0;
            int rstart = rmag.Length - 1;
            for (int i = xlen - 1; i >= 0; i--)
            {
                long product = (value[i] & LONG_MASK) * dl + carry;
                rmag[rstart--] = (int)product;
                carry = Operator.UnsignedRightShift(product, 32);
            }
            rmag[rstart] = (int)carry;
            if (dh != 0L)
            {
                carry = 0;
                rstart = rmag.Length - 2;
                for (int i = xlen - 1; i >= 0; i--)
                {
                    long product = (value[i] & LONG_MASK) * dh +
                        (rmag[rstart] & LONG_MASK) + carry;
                    rmag[rstart--] = (int)product;
                    carry = Operator.UnsignedRightShift(product, 32);
                }
                rmag[0] = (int)carry;
            }
            if (carry == 0L)
                rmag = Arrays.CopyOfRange(rmag, 1, rmag.Length);
            return new BigInteger(rmag, rsign);
        }
        /**
         * Returns a BigInteger whose value is {@code (this * val)}.
         *
         * @param  val value to be multiplied by this BigInteger.
         * @return {@code this * val}
         */
        public BigInteger Multiply(BigInteger val)
        {
            if (val._signum == 0 || _signum == 0)
                return ZERO;

            int[] result = MultiplyToLen(mag, mag.Length,
                                         val.mag, val.mag.Length, null);
            result = TrustedStripLeadingZeroInts(result);
            return new BigInteger(result, _signum == val._signum ? 1 : -1);
        }
        /**
         * Returns a BigInteger whose value is {@code (this + val)}.
         *
         * @param  val value to be added to this BigInteger.
         * @return {@code this + val}
         */
        public BigInteger Add(BigInteger val)
        {
            if (val._signum == 0)
                return this;
            if (_signum == 0)
                return val;
            if (val._signum == _signum)
                return new BigInteger(add(mag, val.mag), _signum);

            int cmp = compareMagnitude(val);
            if (cmp == 0)
                return ZERO;
            int[] resultMag = (cmp > 0 ? Subtract(mag, val.mag)
                               : Subtract(val.mag, mag));
            resultMag = TrustedStripLeadingZeroInts(resultMag);

            return new BigInteger(resultMag, cmp == _signum ? 1 : -1);
        }
        /**
         * Adds the contents of the int arrays x and y. This method allocates
         * a new int array to hold the answer and returns a reference to that
         * array.
         */
        private static int[] add(int[] x, int[] y)
        {
            // If x is shorter, swap the two arrays
            if (x.Length < y.Length)
            {
                int[] tmp = x;
                x = y;
                y = tmp;
            }

            int xIndex = x.Length;
            int yIndex = y.Length;
            int[] result = new int[xIndex];
            long sum = 0;

            // Add common parts of both numbers
            while (yIndex > 0)
            {
                sum = (x[--xIndex] & LONG_MASK) +
                      (y[--yIndex] & LONG_MASK) + Operator.UnsignedRightShift(sum, 32);
                result[xIndex] = (int)sum;
            }

            // Copy remainder of longer number while carry propagation is required
            bool carry = Operator.UnsignedRightShift(sum, 32) != 0;
            while (xIndex > 0 && carry)
                carry = ((result[--xIndex] = x[xIndex] + 1) == 0);

            // Copy remainder of longer number
            while (xIndex > 0)
                result[--xIndex] = x[xIndex];

            // Grow result if necessary
            if (carry)
            {
                int[] bigger = new int[result.Length + 1];
                Array.Copy(result, 0, bigger, 1, result.Length);
                bigger[0] = 0x01;
                return bigger;
            }
            return result;
        }
        /**
         * Returns a BigInteger whose value is {@code (this - val)}.
         *
         * @param  val value to be subtracted from this BigInteger.
         * @return {@code this - val}
         */
        public BigInteger Subtract(BigInteger val)
        {
            if (val._signum == 0)
                return this;
            if (_signum == 0)
                return val.Negate();
            if (val._signum != _signum)
                return new BigInteger(add(mag, val.mag), _signum);

            int cmp = compareMagnitude(val);
            if (cmp == 0)
                return ZERO;
            int[] resultMag = (cmp > 0 ? Subtract(mag, val.mag)
                               : Subtract(val.mag, mag));
            resultMag = TrustedStripLeadingZeroInts(resultMag);

            return new BigInteger(resultMag, cmp == _signum ? 1 : -1);
        }

        /**
         * Subtracts the contents of the second int arrays (little) from the
         * first (big).  The first int array (big) must represent a larger number
         * than the second.  This method allocates the space necessary to hold the
         * answer.
         */
        private static int[] Subtract(int[] big, int[] little)
        {
            int bigIndex = big.Length;
            int[] result = new int[bigIndex];
            int littleIndex = little.Length;
            long difference = 0;

            // Subtract common parts of both numbers
            while (littleIndex > 0)
            {
                difference = (big[--bigIndex] & LONG_MASK) -
                             (little[--littleIndex] & LONG_MASK) +
                             (difference >> 32);
                result[bigIndex] = (int)difference;
            }

            // Subtract remainder of longer number while borrow propagates
            bool borrow = (difference >> 32 != 0);
            while (bigIndex > 0 && borrow)
                borrow = ((result[--bigIndex] = big[bigIndex] - 1) == -1);

            // Copy remainder of longer number
            while (bigIndex > 0)
                result[--bigIndex] = big[bigIndex];

            return result;
        }
        /**
         * Returns a BigInteger whose value is {@code (this / val)}.
         *
         * @param  val value by which this BigInteger is to be divided.
         * @return {@code this / val}
         * @throws ArithmeticException if {@code val} is zero.
         */
        public BigInteger Divide(BigInteger val)
        {
            MutableBigInteger q = new MutableBigInteger(),
                              a = new MutableBigInteger(this.mag),
                              b = new MutableBigInteger(val.mag);

            a.divide(b, q);
            return q.toBigInteger(this._signum == val._signum ? 1 : -1);
        }
        #region operator overload
        //***********************************************************************
        // Overloading of unary >> operators
        //***********************************************************************
        public static BigInteger operator >>(BigInteger bi1, int shiftVal)
        {
            return bi1.ShiftRight(shiftVal);
        }
        //***********************************************************************
        // Overloading of unary << operators
        //***********************************************************************

        public static BigInteger operator <<(BigInteger bi1, int shiftVal)
        {
            return bi1.ShiftLeft(shiftVal);
        }
        //***********************************************************************
        // Overloading of bitwise AND operator
        //***********************************************************************

        public static BigInteger operator &(BigInteger bi1, BigInteger bi2)
        {
            return bi1.and(bi2);
        }
        //***********************************************************************
        // Overloading of bitwise OR operator
        //***********************************************************************
        public static BigInteger operator |(BigInteger bi1, BigInteger bi2)
        {
            return bi1.Or(bi2);
        }
        //***********************************************************************
        // Overloading of bitwise mul operator
        //***********************************************************************
        public static BigInteger operator *(BigInteger bi1, BigInteger bi2)
        {
            return bi1.Multiply(bi2);
        }
        //***********************************************************************
        // Overloading of bitwise add operator
        //***********************************************************************
        public static BigInteger operator +(BigInteger bi1, BigInteger bi2)
        {
            return bi1.Add(bi2);
        }
        //***********************************************************************
        // Overloading of bitwise subtract operator
        //***********************************************************************
        public static BigInteger operator -(BigInteger bi1, BigInteger bi2)
        {
            return bi1.Subtract(bi2);
        }
        public static bool operator <(BigInteger bi1, BigInteger bi2)
        {
            return bi1.CompareTo(bi2) < 0;
        }
        public static bool operator >(BigInteger bi1, BigInteger bi2)
        {
            return bi1.CompareTo(bi2) > 0;
        }
        public static BigInteger operator /(BigInteger bi1, BigInteger bi2)
        {
            return bi1.Divide(bi2);
        }

        public static bool operator ==(BigInteger bi1, BigInteger bi2)
        {
            return bi1.Equals(bi2);
        }
        public static bool operator !=(BigInteger bi1, BigInteger bi2)
        {
            return !(bi1 == bi2);
        }
        #endregion
    }
    internal class MutableBigInteger
    {
        /**
         * Holds the magnitude of this MutableBigInteger in big endian order.
         * The magnitude may start at an offset into the value array, and it may
         * end before the length of the value array.
         */
        int[] _value;

        /**
         * The number of ints of the value array that are currently used
         * to hold the magnitude of this MutableBigInteger. The magnitude starts
         * at an offset and offset + intLen may be less than value.Length.
         */
        int intLen;

        /**
         * The offset into the value array where the magnitude of this
         * MutableBigInteger begins.
         */
        int offset = 0;

        // Constants
        /**
         * MutableBigInteger with one element value array with the value 1. Used by
         * BigDecimal divideAndRound to increment the quotient. Use this constant
         * only when the method is not going to modify this object.
         */
        static readonly MutableBigInteger One = new MutableBigInteger(1);

        // Constructors
        private const long LONG_MASK = BigInteger.LONG_MASK;
        private const long INFLATED = long.MinValue;
        /**
         * The default constructor. An empty MutableBigInteger is created with
         * a one word capacity.
         */
        public MutableBigInteger()
        {
            _value = new int[1];
            intLen = 0;
        }

        /**
         * Construct a new MutableBigInteger with a magnitude specified by
         * the int val.
         */
        public MutableBigInteger(int val)
        {
            _value = new int[1];
            intLen = 1;
            _value[0] = val;
        }

        /**
         * Construct a new MutableBigInteger with the specified value array
         * up to the length of the array supplied.
         */
        public MutableBigInteger(int[] val)
        {
            _value = val;
            intLen = val.Length;
        }
        public static int[] ArraysCopyOf(int[] original, int newLength)
        {
            int[] copy = new int[newLength];
            Array.Copy(original, 0, copy, 0,
                             Math.Min(original.Length, newLength));
            return copy;
        }
        public static long[] ArraysCopyOf(long[] original, int newLength)
        {
            long[] copy = new long[newLength];
            Array.Copy(original, 0, copy, 0,
                             Math.Min(original.Length, newLength));
            return copy;
        }
        public static int[] ArraysCopyOfRange(int[] original, int from, int to)
        {
            int newLength = to - from;
            if (newLength < 0)
                throw new ArgumentException(from + " > " + to);
            int[] copy = new int[newLength];
            Array.Copy(original, from, copy, 0,
                             Math.Min(original.Length - from, newLength));
            return copy;
        }
        public static long[] ArraysCopyOfRange(long[] original, int from, int to)
        {
            int newLength = to - from;
            if (newLength < 0)
                throw new ArgumentException(from + " > " + to);
            long[] copy = new long[newLength];
            Array.Copy(original, from, copy, 0,
                             Math.Min(original.Length - from, newLength));
            return copy;
        }
        /**
         * Construct a new MutableBigInteger with a magnitude equal to the
         * specified BigInteger.
         */
        MutableBigInteger(BigInteger b)
        {
            intLen = b.mag.Length;
            _value = ArraysCopyOf(b.mag, intLen);
        }

        /**
         * Construct a new MutableBigInteger with a magnitude equal to the
         * specified MutableBigInteger.
         */
        MutableBigInteger(MutableBigInteger val)
        {
            intLen = val.intLen;
            _value = ArraysCopyOfRange(val._value, val.offset, val.offset + intLen);
        }

        /**
         * Internal helper method to return the magnitude array. The caller is not
         * supposed to modify the returned array.
         */
        private int[] getMagnitudeArray()
        {
            if (offset > 0 || _value.Length != intLen)
                return ArraysCopyOfRange(_value, offset, offset + intLen);
            return _value;
        }

        /**
         * Convert this MutableBigInteger to a long value. The caller has to make
         * sure this MutableBigInteger can be fit into long.
         */
        private long toLong()
        {
            Debug.Assert(intLen <= 2, "this MutableBigInteger exceeds the range of long");
            if (intLen == 0)
                return 0;
            long d = _value[offset] & LONG_MASK;
            return (intLen == 2) ? d << 32 | (_value[offset + 1] & LONG_MASK) : d;
        }

        /**
         * Convert this MutableBigInteger to a BigInteger object.
         */
        public BigInteger toBigInteger(int sign)
        {
            if (intLen == 0 || sign == 0)
                return BigInteger.ZERO;
            return new BigInteger(getMagnitudeArray(), sign);
        }

        /*
         * Convert this MutableBigInteger to BigDecimal object with the specified sign
         * and scale.
         */
        //BigDecimal toBigDecimal(int sign, int scale) {
        //    if (intLen == 0 || sign == 0)
        //        return BigDecimal.valueOf(0, scale);
        //    int[] mag = getMagnitudeArray();
        //    int len = mag.Length;
        //    int d = mag[0];
        //    // If this MutableBigInteger can't be fit into long, we need to
        //    // make a BigInteger object for the resultant BigDecimal object.
        //    if (len > 2 || (d < 0 && len == 2))
        //        return new BigDecimal(new BigInteger(mag, sign), INFLATED, scale, 0);
        //    long v = (len == 2) ?
        //        ((mag[1] & LONG_MASK) | (d & LONG_MASK) << 32) :
        //        d & LONG_MASK;
        //    return new BigDecimal(null, sign == -1 ? -v : v, scale, 0);
        //}

        /**
         * Clear out a MutableBigInteger for reuse.
         */
        void clear()
        {
            offset = intLen = 0;
            for (int index = 0, n = _value.Length; index < n; index++)
                _value[index] = 0;
        }

        /**
         * Set a MutableBigInteger to zero, removing its offset.
         */
        void reset()
        {
            offset = intLen = 0;
        }

        /**
         * Compare the magnitude of two MutableBigIntegers. Returns -1, 0 or 1
         * as this MutableBigInteger is numerically less than, equal to, or
         * greater than <c>b</c>.
         */
        int compare(MutableBigInteger b)
        {
            int blen = b.intLen;
            if (intLen < blen)
                return -1;
            if (intLen > blen)
                return 1;

            // Add Integer.Min_VALUE to make the comparison act as unsigned integer
            // comparison.
            int[] bval = b._value;
            for (int i = offset, j = b.offset; i < intLen + offset; i++, j++)
            {
                int b1 = unchecked((int)(_value[i] + 0x80000000));
                int b2 = unchecked((int)(bval[j] + 0x80000000));
                if (b1 < b2)
                    return -1;
                if (b1 > b2)
                    return 1;
            }
            return 0;
        }

        /**
         * Compare this against half of a MutableBigInteger object (Needed for
         * remainder tests).
         * Assumes no leading unnecessary zeros, which holds for results
         * from divide().
         */
        int compareHalf(MutableBigInteger b)
        {
            int blen = b.intLen;
            int len = intLen;
            if (len <= 0)
                return blen <= 0 ? 0 : -1;
            if (len > blen)
                return 1;
            if (len < blen - 1)
                return -1;
            int[] bval = b._value;
            int bstart = 0;
            int carry = 0;
            // Only 2 cases left:len == blen or len == blen - 1
            if (len != blen)
            { // len == blen - 1
                if (bval[bstart] == 1)
                {
                    ++bstart;
                    carry = unchecked((int)0x80000000);
                }
                else
                    return -1;
            }
            // compare values with right-shifted values of b,
            // carrying shifted-out bits across words
            int[] val = _value;
            for (int i = offset, j = bstart; i < len + offset; )
            {
                int bv = bval[j++];
                long hb = (Operator.UnsignedRightShift(bv, 1) + carry) & LONG_MASK;
                long v = val[i++] & LONG_MASK;
                if (v != hb)
                    return v < hb ? -1 : 1;
                carry = (bv & 1) << 31; // carray will be either 0x80000000 or 0
            }
            return carry == 0 ? 0 : -1;
        }

        /**
         * Return the index of the lowest set bit in this MutableBigInteger. If the
         * magnitude of this MutableBigInteger is zero, -1 is returned.
         */
        private int getLowestSetBit()
        {
            if (intLen == 0)
                return -1;
            int j, b;
            for (j = intLen - 1; (j > 0) && (_value[j + offset] == 0); j--)
                ;
            b = _value[j + offset];
            if (b == 0)
                return -1;
            return ((intLen - 1 - j) << 5) + BigInteger.NumberOfTrailingZeros(b);
        }

        /**
         * Return the int in use in this MutableBigInteger at the specified
         * index. This method is not used because it is not inlined on all
         * platforms.
         */
        private int getInt(int index)
        {
            return _value[offset + index];
        }

        /**
         * Return a long which is equal to the unsigned value of the int in
         * use in this MutableBigInteger at the specified index. This method is
         * not used because it is not inlined on all platforms.
         */
        private long getLong(int index)
        {
            return _value[offset + index] & LONG_MASK;
        }

        /**
         * Ensure that the MutableBigInteger is in normal form, specifically
         * making sure that there are no leading zeros, and that if the
         * magnitude is zero, then intLen is zero.
         */
        void normalize()
        {
            if (intLen == 0)
            {
                offset = 0;
                return;
            }

            int index = offset;
            if (_value[index] != 0)
                return;

            int indexBound = index + intLen;
            do
            {
                index++;
            } while (index < indexBound && _value[index] == 0);

            int numZeros = index - offset;
            intLen -= numZeros;
            offset = (intLen == 0 ? 0 : offset + numZeros);
        }

        /**
         * If this MutableBigInteger cannot hold len words, increase the size
         * of the value array to len words.
         */
        private void ensureCapacity(int len)
        {
            if (_value.Length < len)
            {
                _value = new int[len];
                offset = 0;
                intLen = len;
            }
        }

        /**
         * Convert this MutableBigInteger into an int array with no leading
         * zeros, of a length that is equal to this MutableBigInteger's intLen.
         */
        int[] toIntArray()
        {
            int[] result = new int[intLen];
            for (int i = 0; i < intLen; i++)
                result[i] = _value[offset + i];
            return result;
        }

        /**
         * Sets the int at index+offset in this MutableBigInteger to val.
         * This does not get inlined on all platforms so it is not used
         * as often as originally intended.
         */
        void setInt(int index, int val)
        {
            _value[offset + index] = val;
        }

        /**
         * Sets this MutableBigInteger's value array to the specified array.
         * The intLen is set to the specified length.
         */
        void setValue(int[] val, int length)
        {
            _value = val;
            intLen = length;
            offset = 0;
        }

        /**
         * Sets this MutableBigInteger's value array to a copy of the specified
         * array. The intLen is set to the length of the new array.
         */
        void copyValue(MutableBigInteger src)
        {
            int len = src.intLen;
            if (_value.Length < len)
                _value = new int[len];
            Array.Copy(src._value, src.offset, _value, 0, len);
            intLen = len;
            offset = 0;
        }

        /**
         * Sets this MutableBigInteger's value array to a copy of the specified
         * array. The intLen is set to the length of the specified array.
         */
        void copyValue(int[] val)
        {
            int len = val.Length;
            if (_value.Length < len)
                _value = new int[len];
            Array.Copy(val, 0, _value, 0, len);
            intLen = len;
            offset = 0;
        }

        /**
         * Returns true iff this MutableBigInteger has a value of one.
         */
        bool isOne()
        {
            return (intLen == 1) && (_value[offset] == 1);
        }

        /**
         * Returns true iff this MutableBigInteger has a value of zero.
         */
        bool isZero()
        {
            return (intLen == 0);
        }

        /**
         * Returns true iff this MutableBigInteger is even.
         */
        bool isEven()
        {
            return (intLen == 0) || ((_value[offset + intLen - 1] & 1) == 0);
        }

        /**
         * Returns true iff this MutableBigInteger is odd.
         */
        bool isOdd()
        {
            return isZero() ? false : ((_value[offset + intLen - 1] & 1) == 1);
        }

        /**
         * Returns true iff this MutableBigInteger is in normal form. A
         * MutableBigInteger is in normal form if it has no leading zeros
         * after the offset, and intLen + offset &lt;= value.Length.
         */
        bool isNormal()
        {
            if (intLen + offset > _value.Length)
                return false;
            if (intLen == 0)
                return true;
            return (_value[offset] != 0);
        }

        /**
         * Returns a String representation of this MutableBigInteger in radix 10.
         */
        public String toString()
        {
            BigInteger b = toBigInteger(1);
            return b.ToString();
        }

        /**
         * Right shift this MutableBigInteger n bits. The MutableBigInteger is left
         * in normal form.
         */
        void rightShift(int n)
        {
            if (intLen == 0)
                return;
            int nInts = Operator.UnsignedRightShift(n, 5);
            int nBits = n & 0x1F;
            this.intLen -= nInts;
            if (nBits == 0)
                return;
            int bitsInHighWord = BigInteger.BitLengthForInt(_value[offset]);
            if (nBits >= bitsInHighWord)
            {
                this.primitiveLeftShift(32 - nBits);
                this.intLen--;
            }
            else
            {
                primitiveRightShift(nBits);
            }
        }

        /**
         * Left shift this MutableBigInteger n bits.
         */
        void leftShift(int n)
        {
            /*
             * If there is enough storage space in this MutableBigInteger already
             * the available space will be used. Space to the right of the used
             * ints in the value array is faster to utilize, so the extra space
             * will be taken from the right if possible.
             */
            if (intLen == 0)
                return;
            int nInts = Operator.UnsignedRightShift(n, 5);
            int nBits = n & 0x1F;
            int bitsInHighWord = BigInteger.BitLengthForInt(_value[offset]);

            // If shift can be done without moving words, do so
            if (n <= (32 - bitsInHighWord))
            {
                primitiveLeftShift(nBits);
                return;
            }

            int newLen = intLen + nInts + 1;
            if (nBits <= (32 - bitsInHighWord))
                newLen--;
            if (_value.Length < newLen)
            {
                // The array must grow
                int[] result = new int[newLen];
                for (int i = 0; i < intLen; i++)
                    result[i] = _value[offset + i];
                setValue(result, newLen);
            }
            else if (_value.Length - offset >= newLen)
            {
                // Use space on right
                for (int i = 0; i < newLen - intLen; i++)
                    _value[offset + intLen + i] = 0;
            }
            else
            {
                // Must use space on left
                for (int i = 0; i < intLen; i++)
                    _value[i] = _value[offset + i];
                for (int i = intLen; i < newLen; i++)
                    _value[i] = 0;
                offset = 0;
            }
            intLen = newLen;
            if (nBits == 0)
                return;
            if (nBits <= (32 - bitsInHighWord))
                primitiveLeftShift(nBits);
            else
                primitiveRightShift(32 - nBits);
        }

        /**
         * A primitive used for division. This method adds in one multiple of the
         * divisor a back to the dividend result at a specified offset. It is used
         * when qhat was estimated too large, and must be adjusted.
         */
        private int divadd(int[] a, int[] result, int offset)
        {
            long carry = 0;

            for (int j = a.Length - 1; j >= 0; j--)
            {
                long sum = (a[j] & LONG_MASK) +
                           (result[j + offset] & LONG_MASK) + carry;
                result[j + offset] = (int)sum;
                carry = Operator.UnsignedRightShift(sum, 32);
            }
            return (int)carry;
        }

        /**
         * This method is used for division. It multiplies an n word input a by one
         * word input x, and subtracts the n word product from q. This is needed
         * when subtracting qhat*divisor from dividend.
         */
        private int mulsub(int[] q, int[] a, int x, int len, int offset)
        {
            long xLong = x & LONG_MASK;
            long carry = 0;
            offset += len;

            for (int j = len - 1; j >= 0; j--)
            {
                long product = (a[j] & LONG_MASK) * xLong + carry;
                long difference = q[offset] - product;
                q[offset--] = (int)difference;
                carry = Operator.UnsignedRightShift(product, 32)
                         + (((difference & LONG_MASK) >
                             (((~(int)product) & LONG_MASK))) ? 1 : 0);
            }
            return (int)carry;
        }

        /**
         * Right shift this MutableBigInteger n bits, where n is
         * less than 32.
         * Assumes that intLen > 0, n > 0 for speed
         */
        private void primitiveRightShift(int n)
        {
            int[] val = _value;
            int n2 = 32 - n;
            for (int i = offset + intLen - 1, c = val[i]; i > offset; i--)
            {
                int b = c;
                c = val[i - 1];
                val[i] = (c << n2) | Operator.UnsignedRightShift(b, n);
            }
            val[offset] = Operator.UnsignedRightShift(val[offset], n);
        }

        /**
         * Left shift this MutableBigInteger n bits, where n is
         * less than 32.
         * Assumes that intLen > 0, n > 0 for speed
         */
        private void primitiveLeftShift(int n)
        {
            int[] val = _value;
            int n2 = 32 - n;
            for (int i = offset, c = val[i], m = i + intLen - 1; i < m; i++)
            {
                int b = c;
                c = val[i + 1];
                val[i] = (b << n) | Operator.UnsignedRightShift(c, n2);
            }
            val[offset + intLen - 1] <<= n;
        }

        /**
         * Adds the contents of two MutableBigInteger objects.The result
         * is placed within this MutableBigInteger.
         * The contents of the addend are not changed.
         */
        void add(MutableBigInteger addend)
        {
            int x = intLen;
            int y = addend.intLen;
            int resultLen = (intLen > addend.intLen ? intLen : addend.intLen);
            int[] result = (_value.Length < resultLen ? new int[resultLen] : _value);

            int rstart = result.Length - 1;
            long sum;
            long carry = 0;

            // Add common parts of both numbers
            while (x > 0 && y > 0)
            {
                x--; y--;
                sum = (_value[x + offset] & LONG_MASK) +
                    (addend._value[y + addend.offset] & LONG_MASK) + carry;
                result[rstart--] = (int)sum;
                carry = Operator.UnsignedRightShift(sum, 32);
            }

            // Add remainder of the longer number
            while (x > 0)
            {
                x--;
                if (carry == 0 && result == _value && rstart == (x + offset))
                    return;
                sum = (_value[x + offset] & LONG_MASK) + carry;
                result[rstart--] = (int)sum;
                carry = Operator.UnsignedRightShift(sum, 32);
            }
            while (y > 0)
            {
                y--;
                sum = (addend._value[y + addend.offset] & LONG_MASK) + carry;
                result[rstart--] = (int)sum;
                carry = Operator.UnsignedRightShift(sum, 32);
            }

            if (carry > 0)
            { // Result must grow in length
                resultLen++;
                if (result.Length < resultLen)
                {
                    int[] temp = new int[resultLen];
                    // Result one word longer from carry-out; copy low-order
                    // bits into new result.
                    Array.Copy(result, 0, temp, 1, result.Length);
                    temp[0] = 1;
                    result = temp;
                }
                else
                {
                    result[rstart--] = 1;
                }
            }

            _value = result;
            intLen = resultLen;
            offset = result.Length - resultLen;
        }


        /**
         * Subtracts the smaller of this and b from the larger and places the
         * result into this MutableBigInteger.
         */
        int subtract(MutableBigInteger b)
        {
            MutableBigInteger a = this;

            int[] result = _value;
            int sign = a.compare(b);

            if (sign == 0)
            {
                reset();
                return 0;
            }
            if (sign < 0)
            {
                MutableBigInteger tmp = a;
                a = b;
                b = tmp;
            }

            int resultLen = a.intLen;
            if (result.Length < resultLen)
                result = new int[resultLen];

            long diff = 0;
            int x = a.intLen;
            int y = b.intLen;
            int rstart = result.Length - 1;

            // Subtract common parts of both numbers
            while (y > 0)
            {
                x--; y--;

                diff = (a._value[x + a.offset] & LONG_MASK) -
                       (b._value[y + b.offset] & LONG_MASK) - ((int)-(diff >> 32));
                result[rstart--] = (int)diff;
            }
            // Subtract remainder of longer number
            while (x > 0)
            {
                x--;
                diff = (a._value[x + a.offset] & LONG_MASK) - ((int)-(diff >> 32));
                result[rstart--] = (int)diff;
            }

            _value = result;
            intLen = resultLen;
            offset = _value.Length - resultLen;
            normalize();
            return sign;
        }

        /**
         * Subtracts the smaller of a and b from the larger and places the result
         * into the larger. Returns 1 if the answer is in a, -1 if in b, 0 if no
         * operation was performed.
         */
        private int difference(MutableBigInteger b)
        {
            MutableBigInteger a = this;
            int sign = a.compare(b);
            if (sign == 0)
                return 0;
            if (sign < 0)
            {
                MutableBigInteger tmp = a;
                a = b;
                b = tmp;
            }

            long diff = 0;
            int x = a.intLen;
            int y = b.intLen;

            // Subtract common parts of both numbers
            while (y > 0)
            {
                x--; y--;
                diff = (a._value[a.offset + x] & LONG_MASK) -
                    (b._value[b.offset + y] & LONG_MASK) - ((int)-(diff >> 32));
                a._value[a.offset + x] = (int)diff;
            }
            // Subtract remainder of longer number
            while (x > 0)
            {
                x--;
                diff = (a._value[a.offset + x] & LONG_MASK) - ((int)-(diff >> 32));
                a._value[a.offset + x] = (int)diff;
            }

            a.normalize();
            return sign;
        }

        /**
         * Multiply the contents of two MutableBigInteger objects. The result is
         * placed into MutableBigInteger z. The contents of y are not changed.
         */
        void multiply(MutableBigInteger y, MutableBigInteger z)
        {
            int xLen = intLen;
            int yLen = y.intLen;
            int newLen = xLen + yLen;

            // Put z into an appropriate state to receive product
            if (z._value.Length < newLen)
                z._value = new int[newLen];
            z.offset = 0;
            z.intLen = newLen;

            // The first iteration is hoisted out of the loop to avoid extra add
            long carry = 0;
            for (int j = yLen - 1, k = yLen + xLen - 1; j >= 0; j--, k--)
            {
                long product = (y._value[j + y.offset] & LONG_MASK) *
                               (_value[xLen - 1 + offset] & LONG_MASK) + carry;
                z._value[k] = (int)product;
                carry = Operator.UnsignedRightShift(product, 32);
            }
            z._value[xLen - 1] = (int)carry;

            // Perform the multiplication word by word
            for (int i = xLen - 2; i >= 0; i--)
            {
                carry = 0;
                for (int j = yLen - 1, k = yLen + i; j >= 0; j--, k--)
                {
                    long product = (y._value[j + y.offset] & LONG_MASK) *
                                   (_value[i + offset] & LONG_MASK) +
                                   (z._value[k] & LONG_MASK) + carry;
                    z._value[k] = (int)product;
                    carry = Operator.UnsignedRightShift(product, 32);
                }
                z._value[i] = (int)carry;
            }

            // Remove leading zeros from product
            z.normalize();
        }

        /**
         * Multiply the contents of this MutableBigInteger by the word y. The
         * result is placed into z.
         */
        public void mul(int y, MutableBigInteger z)
        {
            if (y == 1)
            {
                z.copyValue(this);
                return;
            }

            if (y == 0)
            {
                z.clear();
                return;
            }

            // Perform the multiplication word by word
            long ylong = y & LONG_MASK;
            int[] zval = (z._value.Length < intLen + 1 ? new int[intLen + 1]
                                                  : z._value);
            long carry = 0;
            for (int i = intLen - 1; i >= 0; i--)
            {
                long product = ylong * (_value[i + offset] & LONG_MASK) + carry;
                zval[i + 1] = (int)product;
                carry = Operator.UnsignedRightShift(product, 32);
            }

            if (carry == 0)
            {
                z.offset = 1;
                z.intLen = intLen;
            }
            else
            {
                z.offset = 0;
                z.intLen = intLen + 1;
                zval[0] = (int)carry;
            }
            z._value = zval;
        }

        /**
        * This method is used for division of an n word dividend by a one word
        * divisor. The quotient is placed into quotient. The one word divisor is
        * specified by divisor.
        *
        * @return the remainder of the division is returned.
        *
        */
        public int divideOneWord(int divisor, MutableBigInteger quotient)
        {
            long divisorLong = divisor & LONG_MASK;

            // Special case of one word dividend
            if (intLen == 1)
            {
                long dividendValue = _value[offset] & LONG_MASK;
                int q = (int)(dividendValue / divisorLong);
                int r = (int)(dividendValue - q * divisorLong);
                quotient._value[0] = q;
                quotient.intLen = (q == 0) ? 0 : 1;
                quotient.offset = 0;
                return r;
            }

            if (quotient._value.Length < intLen)
                quotient._value = new int[intLen];
            quotient.offset = 0;
            quotient.intLen = intLen;

            // Normalize the divisor
            int shift = BigInteger.NumberOfLeadingZeros(divisor);

            int rem = _value[offset];
            long remLong = rem & LONG_MASK;
            if (remLong < divisorLong)
            {
                quotient._value[0] = 0;
            }
            else
            {
                quotient._value[0] = (int)(remLong / divisorLong);
                rem = (int)(remLong - (quotient._value[0] * divisorLong));
                remLong = rem & LONG_MASK;
            }

            int xlen = intLen;
            int[] qWord = new int[2];
            while (--xlen > 0)
            {
                long dividendEstimate = (remLong << 32) |
                    (_value[offset + intLen - xlen] & LONG_MASK);
                if (dividendEstimate >= 0)
                {
                    qWord[0] = (int)(dividendEstimate / divisorLong);
                    qWord[1] = (int)(dividendEstimate - qWord[0] * divisorLong);
                }
                else
                {
                    divWord(qWord, dividendEstimate, divisor);
                }
                quotient._value[intLen - xlen] = qWord[0];
                rem = qWord[1];
                remLong = rem & LONG_MASK;
            }

            quotient.normalize();
            // Unnormalize
            if (shift > 0)
                return rem % divisor;
            else
                return rem;
        }

        /**
         * Calculates the quotient of this div b and places the quotient in the
         * provided MutableBigInteger objects and the remainder object is returned.
         *
         * Uses Algorithm D in Knuth section 4.3.1.
         * Many optimizations to that algorithm have been adapted from the Colin
         * Plumb C library.
         * It special cases one word divisors for speed. The content of b is not
         * changed.
         *
         */
        public MutableBigInteger divide(MutableBigInteger b, MutableBigInteger quotient)
        {
            if (b.intLen == 0)
                throw new ArithmeticException("BigInteger divide by zero");

            // Dividend is zero
            if (intLen == 0)
            {
                quotient.intLen = quotient.offset;
                return new MutableBigInteger();
            }

            int cmp = compare(b);
            // Dividend less than divisor
            if (cmp < 0)
            {
                quotient.intLen = quotient.offset = 0;
                return new MutableBigInteger(this);
            }
            // Dividend equal to divisor
            if (cmp == 0)
            {
                quotient._value[0] = quotient.intLen = 1;
                quotient.offset = 0;
                return new MutableBigInteger();
            }

            quotient.clear();
            // Special case one word divisor
            if (b.intLen == 1)
            {
                int r = divideOneWord(b._value[b.offset], quotient);
                if (r == 0)
                    return new MutableBigInteger();
                return new MutableBigInteger(r);
            }

            // Copy divisor value to protect divisor
            int[] div = ArraysCopyOfRange(b._value, b.offset, b.offset + b.intLen);
            return divideMagnitude(div, quotient);
        }

        /**
         * Internally used  to calculate the quotient of this div v and places the
         * quotient in the provided MutableBigInteger object and the remainder is
         * returned.
         *
         * @return the remainder of the division will be returned.
         */
        public long divide(long v, MutableBigInteger quotient)
        {
            if (v == 0)
                throw new ArithmeticException("BigInteger divide by zero");

            // Dividend is zero
            if (intLen == 0)
            {
                quotient.intLen = quotient.offset = 0;
                return 0;
            }
            if (v < 0)
                v = -v;

            int d = (int)Operator.UnsignedRightShift(v, 32);
            quotient.clear();
            // Special case on word divisor
            if (d == 0)
                return divideOneWord((int)v, quotient) & LONG_MASK;
            else
            {
                int[] div = new int[] { d, (int)(v & LONG_MASK) };
                return divideMagnitude(div, quotient).toLong();
            }
        }

        /**
         * Divide this MutableBigInteger by the divisor represented by its magnitude
         * array. The quotient will be placed into the provided quotient object &amp;
         * the remainder object is returned.
         */
        private MutableBigInteger divideMagnitude(int[] divisor,
                                                  MutableBigInteger quotient)
        {

            // Remainder starts as dividend with space for a leading zero
            MutableBigInteger rem = new MutableBigInteger(new int[intLen + 1]);
            Array.Copy(_value, offset, rem._value, 1, intLen);
            rem.intLen = intLen;
            rem.offset = 1;

            int nlen = rem.intLen;

            // Set the quotient size
            int dlen = divisor.Length;
            int limit = nlen - dlen + 1;
            if (quotient._value.Length < limit)
            {
                quotient._value = new int[limit];
                quotient.offset = 0;
            }
            quotient.intLen = limit;
            int[] q = quotient._value;

            // D1 normalize the divisor
            int shift = BigInteger.NumberOfLeadingZeros(divisor[0]);
            if (shift > 0)
            {
                // First shift will not grow array
                BigInteger.PrimitiveLeftShift(divisor, dlen, shift);
                // But this one might
                rem.leftShift(shift);
            }

            // Must insert leading 0 in rem if its length did not change
            if (rem.intLen == nlen)
            {
                rem.offset = 0;
                rem._value[0] = 0;
                rem.intLen++;
            }

            int dh = divisor[0];
            long dhLong = dh & LONG_MASK;
            int dl = divisor[1];
            int[] qWord = new int[2];

            // D2 Initialize j
            for (int j = 0; j < limit; j++)
            {
                // D3 Calculate qhat
                // estimate qhat
                int qhat = 0;
                int qrem = 0;
                bool skipCorrection = false;
                int nh = rem._value[j + rem.offset];
                int nh2 = unchecked((int)(nh + 0x80000000));
                int nm = rem._value[j + 1 + rem.offset];

                if (nh == dh)
                {
                    qhat = ~0;
                    qrem = nh + nm;
                    skipCorrection = qrem + 0x80000000 < nh2;
                }
                else
                {
                    long nChunk = (((long)nh) << 32) | (nm & LONG_MASK);
                    if (nChunk >= 0)
                    {
                        qhat = (int)(nChunk / dhLong);
                        qrem = (int)(nChunk - (qhat * dhLong));
                    }
                    else
                    {
                        divWord(qWord, nChunk, dh);
                        qhat = qWord[0];
                        qrem = qWord[1];
                    }
                }

                if (qhat == 0)
                    continue;

                if (!skipCorrection)
                { // Correct qhat
                    long nl = rem._value[j + 2 + rem.offset] & LONG_MASK;
                    long rs = ((qrem & LONG_MASK) << 32) | nl;
                    long estProduct = (dl & LONG_MASK) * (qhat & LONG_MASK);

                    if (unsignedLongCompare(estProduct, rs))
                    {
                        qhat--;
                        qrem = (int)((qrem & LONG_MASK) + dhLong);
                        if ((qrem & LONG_MASK) >= dhLong)
                        {
                            estProduct -= (dl & LONG_MASK);
                            rs = ((qrem & LONG_MASK) << 32) | nl;
                            if (unsignedLongCompare(estProduct, rs))
                                qhat--;
                        }
                    }
                }

                // D4 Multiply and subtract
                rem._value[j + rem.offset] = 0;
                int borrow = mulsub(rem._value, divisor, qhat, dlen, j + rem.offset);

                // D5 Test remainder
                if ((int)(borrow + 0x80000000) > nh2)
                {
                    // D6 Add back
                    divadd(divisor, rem._value, j + 1 + rem.offset);
                    qhat--;
                }

                // Store the quotient digit
                q[j] = qhat;
            } // D7 loop on j

            // D8 Unnormalize
            if (shift > 0)
                rem.rightShift(shift);

            quotient.normalize();
            rem.normalize();
            return rem;
        }

        /**
         * Compare two longs as if they were unsigned.
         * Returns true iff one is bigger than two.
         */
        private bool unsignedLongCompare(long one, long two)
        {
            return (one + long.MinValue) > (two + long.MinValue);
        }

        /**
         * This method divides a long quantity by an int to estimate
         * qhat for two multi precision numbers. It is used when
         * the signed value of n is less than zero.
         */
        private void divWord(int[] result, long n, int d)
        {
            long dLong = d & LONG_MASK;

            if (dLong == 1)
            {
                result[0] = (int)n;
                result[1] = 0;
                return;
            }

            // Approximate the quotient and remainder
            long q = Operator.UnsignedRightShift(n, 1) / Operator.UnsignedRightShift(dLong, 1);
            long r = n - q * dLong;

            // Correct the approximation
            while (r < 0)
            {
                r += dLong;
                q--;
            }
            while (r >= dLong)
            {
                r -= dLong;
                q++;
            }

            // n - q*dlong == r && 0 <= r <dLong, hence we're done.
            result[0] = (int)q;
            result[1] = (int)r;
        }

        /**
         * Calculate GCD of this and b. This and b are changed by the computation.
         */
        MutableBigInteger hybridGCD(MutableBigInteger b)
        {
            // Use Euclid's algorithm until the numbers are approximately the
            // same length, then use the binary GCD algorithm to find the GCD.
            MutableBigInteger a = this;
            MutableBigInteger q = new MutableBigInteger();

            while (b.intLen != 0)
            {
                if (Math.Abs(a.intLen - b.intLen) < 2)
                    return a.binaryGCD(b);

                MutableBigInteger r = a.divide(b, q);
                a = b;
                b = r;
            }
            return a;
        }

        /**
         * Calculate GCD of this and v.
         * Assumes that this and v are not zero.
         */
        private MutableBigInteger binaryGCD(MutableBigInteger v)
        {
            // Algorithm B from Knuth section 4.5.2
            MutableBigInteger u = this;
            MutableBigInteger r = new MutableBigInteger();

            // step B1
            int s1 = u.getLowestSetBit();
            int s2 = v.getLowestSetBit();
            int k = (s1 < s2) ? s1 : s2;
            if (k != 0)
            {
                u.rightShift(k);
                v.rightShift(k);
            }

            // step B2
            bool uOdd = (k == s1);
            MutableBigInteger t = uOdd ? v : u;
            int tsign = uOdd ? -1 : 1;

            int lb;
            while ((lb = t.getLowestSetBit()) >= 0)
            {
                // steps B3 and B4
                t.rightShift(lb);
                // step B5
                if (tsign > 0)
                    u = t;
                else
                    v = t;

                // Special case one word numbers
                if (u.intLen < 2 && v.intLen < 2)
                {
                    int x = u._value[u.offset];
                    int y = v._value[v.offset];
                    x = binaryGcd(x, y);
                    r._value[0] = x;
                    r.intLen = 1;
                    r.offset = 0;
                    if (k > 0)
                        r.leftShift(k);
                    return r;
                }

                // step B6
                if ((tsign = u.difference(v)) == 0)
                    break;
                t = (tsign >= 0) ? u : v;
            }

            if (k > 0)
                u.leftShift(k);
            return u;
        }

        /**
         * Calculate GCD of a and b interpreted as unsigned integers.
         */
        static int binaryGcd(int a, int b)
        {
            if (b == 0)
                return a;
            if (a == 0)
                return b;

            // Right shift a & b till their last bits equal to 1.
            int aZeros = BigInteger.NumberOfTrailingZeros(a);
            int bZeros = BigInteger.NumberOfTrailingZeros(b);
            a = Operator.UnsignedRightShift(a, aZeros);
            b = Operator.UnsignedRightShift(b, bZeros);

            int t = (aZeros < bZeros ? aZeros : bZeros);

            while (a != b)
            {
                if ((a + 0x80000000) > (b + 0x80000000))
                {  // a > b as unsigned
                    a -= b;
                    a = Operator.UnsignedRightShift(a, BigInteger.NumberOfTrailingZeros(a));
                }
                else
                {
                    b -= a;
                    b = Operator.UnsignedRightShift(b, BigInteger.NumberOfTrailingZeros(b));
                }
            }
            return a << t;
        }

        /**
         * Returns the modInverse of this mod p. This and p are not affected by
         * the operation.
         */
        MutableBigInteger mutableModInverse(MutableBigInteger p)
        {
            // Modulus is odd, use Schroeppel's algorithm
            if (p.isOdd())
                return modInverse(p);

            // Base and modulus are even, throw exception
            if (isEven())
                throw new ArithmeticException("BigInteger not invertible.");

            // Get even part of modulus expressed as a power of 2
            int powersOf2 = p.getLowestSetBit();

            // Construct odd part of modulus
            MutableBigInteger oddMod = new MutableBigInteger(p);
            oddMod.rightShift(powersOf2);

            if (oddMod.isOne())
                return modInverseMP2(powersOf2);

            // Calculate 1/a mod oddMod
            MutableBigInteger oddPart = modInverse(oddMod);

            // Calculate 1/a mod evenMod
            MutableBigInteger evenPart = modInverseMP2(powersOf2);

            // Combine the results using Chinese Remainder Theorem
            MutableBigInteger y1 = modInverseBP2(oddMod, powersOf2);
            MutableBigInteger y2 = oddMod.modInverseMP2(powersOf2);

            MutableBigInteger temp1 = new MutableBigInteger();
            MutableBigInteger temp2 = new MutableBigInteger();
            MutableBigInteger result = new MutableBigInteger();

            oddPart.leftShift(powersOf2);
            oddPart.multiply(y1, result);

            evenPart.multiply(oddMod, temp1);
            temp1.multiply(y2, temp2);

            result.add(temp2);
            return result.divide(p, temp1);
        }

        /*
         * Calculate the multiplicative inverse of this mod 2^k.
         */
        MutableBigInteger modInverseMP2(int k)
        {
            if (isEven())
                throw new ArithmeticException("Non-invertible. (GCD != 1)");

            if (k > 64)
                return euclidModInverse(k);

            int t = inverseMod32(_value[offset + intLen - 1]);

            if (k < 33)
            {
                t = (k == 32 ? t : t & ((1 << k) - 1));
                return new MutableBigInteger(t);
            }

            long pLong = (_value[offset + intLen - 1] & LONG_MASK);
            if (intLen > 1)
                pLong |= ((long)_value[offset + intLen - 2] << 32);
            long tLong = t & LONG_MASK;
            tLong = tLong * (2 - pLong * tLong);  // 1 more Newton iter step
            tLong = (k == 64 ? tLong : tLong & ((1L << k) - 1));

            MutableBigInteger result = new MutableBigInteger(new int[2]);
            result._value[0] = (int)Operator.UnsignedRightShift(tLong, 32);
            result._value[1] = (int)tLong;
            result.intLen = 2;
            result.normalize();
            return result;
        }

        /*
         * Returns the multiplicative inverse of val mod 2^32.  Assumes val is odd.
         */
        static int inverseMod32(int val)
        {
            // Newton's iteration!
            int t = val;
            t *= 2 - val * t;
            t *= 2 - val * t;
            t *= 2 - val * t;
            t *= 2 - val * t;
            return t;
        }

        /*
         * Calculate the multiplicative inverse of 2^k mod mod, where mod is odd.
         */
        static MutableBigInteger modInverseBP2(MutableBigInteger mod, int k)
        {
            // Copy the mod to protect original
            return fixup(new MutableBigInteger(1), new MutableBigInteger(mod), k);
        }

        /**
         * Calculate the multiplicative inverse of this mod mod, where mod is odd.
         * This and mod are not changed by the calculation.
         *
         * This method implements an algorithm due to Richard Schroeppel, that uses
         * the same intermediate representation as Montgomery Reduction
         * ("Montgomery Form").  The algorithm is described in an unpublished
         * manuscript entitled "Fast Modular Reciprocals."
         */
        private MutableBigInteger modInverse(MutableBigInteger mod)
        {
            //TODO: to complete this method,we should implement SignedMutableBigInteger class
            throw new NotImplementedException("This method uses SignedMutableBigInteger class.");
            /*MutableBigInteger p = new MutableBigInteger(mod);
            MutableBigInteger f = new MutableBigInteger(this);
            MutableBigInteger g = new MutableBigInteger(p);
            SignedMutableBigInteger c = new SignedMutableBigInteger(1);
            SignedMutableBigInteger d = new SignedMutableBigInteger();
            MutableBigInteger temp = null;
            SignedMutableBigInteger sTemp = null;

            int k = 0;
            // Right shift f k times until odd, left shift d k times
            if (f.isEven()) {
                int trailingZeros = f.getLowestSetBit();
                f.rightShift(trailingZeros);
                d.leftShift(trailingZeros);
                k = trailingZeros;
            }

            // The Almost Inverse Algorithm
            while(!f.isOne()) {
                // If gcd(f, g) != 1, number is not invertible modulo mod
                if (f.isZero())
                    throw new ArithmeticException("BigInteger not invertible.");

                // If f < g exchange f, g and c, d
                if (f.compare(g) < 0) {
                    temp = f; f = g; g = temp;
                    sTemp = d; d = c; c = sTemp;
                }

                // If f == g (mod 4)
                if (((f.value[f.offset + f.intLen - 1] ^
                     g.value[g.offset + g.intLen - 1]) & 3) == 0) {
                    f.subtract(g);
                    c.signedSubtract(d);
                } else { // If f != g (mod 4)
                    f.add(g);
                    c.signedAdd(d);
                }

                // Right shift f k times until odd, left shift d k times
                int trailingZeros = f.getLowestSetBit();
                f.rightShift(trailingZeros);
                d.leftShift(trailingZeros);
                k += trailingZeros;
            }

            while (c.sign < 0)
               c.signedAdd(p);

            return fixup(c, p, k);*/
        }

        /*
         * The Fixup Algorithm
         * Calculates X such that X = C * 2^(-k) (mod P)
         * Assumes C<P and P is odd.
         */
        static MutableBigInteger fixup(MutableBigInteger c, MutableBigInteger p,
                                                                          int k)
        {
            MutableBigInteger temp = new MutableBigInteger();
            // Set r to the multiplicative inverse of p mod 2^32
            int r = -inverseMod32(p._value[p.offset + p.intLen - 1]);

            for (int i = 0, numWords = k >> 5; i < numWords; i++)
            {
                // V = R * c (mod 2^j)
                int v = r * c._value[c.offset + c.intLen - 1];
                // c = c + (v * p)
                p.mul(v, temp);
                c.add(temp);
                // c = c / 2^j
                c.intLen--;
            }
            int numBits = k & 0x1f;
            if (numBits != 0)
            {
                // V = R * c (mod 2^j)
                int v = r * c._value[c.offset + c.intLen - 1];
                v &= ((1 << numBits) - 1);
                // c = c + (v * p)
                p.mul(v, temp);
                c.add(temp);
                // c = c / 2^j
                c.rightShift(numBits);
            }

            // In theory, c may be greater than p at this point (Very rare!)
            while (c.compare(p) >= 0)
                c.subtract(p);

            return c;
        }

        /**
         * Uses the extended Euclidean algorithm to compute the modInverse of base
         * mod a modulus that is a power of 2. The modulus is 2^k.
         */
        MutableBigInteger euclidModInverse(int k)
        {
            MutableBigInteger b = new MutableBigInteger(1);
            b.leftShift(k);
            MutableBigInteger mod = new MutableBigInteger(b);

            MutableBigInteger a = new MutableBigInteger(this);
            MutableBigInteger q = new MutableBigInteger();
            MutableBigInteger r = b.divide(a, q);

            MutableBigInteger swapper = b;
            // swap b & r
            b = r;
            r = swapper;

            MutableBigInteger t1 = new MutableBigInteger(q);
            MutableBigInteger t0 = new MutableBigInteger(1);
            MutableBigInteger temp = new MutableBigInteger();

            while (!b.isOne())
            {
                r = a.divide(b, q);

                if (r.intLen == 0)
                    throw new ArithmeticException("BigInteger not invertible.");

                swapper = r;
                a = swapper;

                if (q.intLen == 1)
                    t1.mul(q._value[q.offset], temp);
                else
                    q.multiply(t1, temp);
                swapper = q;
                q = temp;
                temp = swapper;
                t0.add(q);

                if (a.isOne())
                    return t0;

                r = b.divide(a, q);

                if (r.intLen == 0)
                    throw new ArithmeticException("BigInteger not invertible.");

                swapper = b;
                b = r;

                if (q.intLen == 1)
                    t0.mul(q._value[q.offset], temp);
                else
                    q.multiply(t0, temp);
                swapper = q; q = temp; temp = swapper;

                t1.add(q);
            }
            mod.subtract(t1);
            return mod;
        }

    }
}
