using System;

namespace NPOI.Util
{
    public class ByteBuffer
    {
        private byte[] buffer;
        private int mark = -1; //useless.
        private int position = 0;
        private int limit;
        private int capacity;
        private int offset;  //not sure about this.

        private ByteBuffer(int mark, int pos, int lim, int cap, byte[] buffer, int offset)
        {
            if (cap < 0)
                throw new ArgumentException();
            this.capacity = cap;

            this.Limit = lim;
            this.Position = pos;
            if (mark >= 0)
            {
                if (mark > pos)
                    throw new ArgumentException();
                this.mark = mark;
            }

            this.buffer = buffer;
            this.offset = offset;
        }

        public ByteBuffer(byte[] buffer, int off, int len)
            :this(-1, off, off+len, buffer.Length, buffer, 0)
        {
        }

        public ByteBuffer(int capacity, int limit)
            :this(-1, 0, limit, capacity, new byte[capacity], 0)
        {
        }

        protected ByteBuffer(byte[] buffer, int mark, int pos, int lim, int cap, int off)
            : this(mark, pos, lim, cap, buffer, off)
        {
        }

        public int Position
        {
            get { return position; }
            set
            {
                if (value < 0 || value > limit)
                    throw new ArgumentException();
                position = value;
                if (mark > position)
                    mark = -1;
            }
        }

        public int Limit
        {
            get { return limit; }
            set
            {
                if ((value > capacity) || (value < 0))
                    throw new ArgumentException();
                limit = value;

                if (position > limit)
                    position = limit;
                if (mark > limit)
                    mark = -1;
            }
        }

        /// <summary>
        /// Returns the number of elements between the current position and the limit.
        /// </summary>
        /// <returns>The number of elements remaining in this buffer</returns>
        public int Remaining() {
            return limit - position;
        }

        /// <summary>
        /// Tells whether there are any elements between the current position and the limit.
        /// </summary>
        /// <returns>true if, and only if, there is at least one element remaining in this buffer</returns>
        public bool HasRemaining() {
            return position < limit;
        }

        //allocate
        public static ByteBuffer CreateBuffer(int capacity)
        {
            if (capacity < 0)
                throw new ArgumentException();
            return new ByteBuffer(capacity, capacity);
        }

        //wrap
        public static ByteBuffer CreateBuffer(byte[] buffer, int offset, int length)
        {
            try
            {
                return new ByteBuffer(buffer, offset, length);
            }
            catch (ArgumentException)
            {
                throw new IndexOutOfRangeException();
            }

        }

        public static ByteBuffer CreateBuffer(byte[] buffer)
        {
            return CreateBuffer(buffer, 0, buffer.Length);
        }

        public ByteBuffer Slice()  //This is not correct! Leon
        {
            return new ByteBuffer(buffer, -1, 0, this.Remain, this.Remain, this.Position + offset);
        }

        public ByteBuffer Duplicate()
        {
            return null;
        }

        protected int NextGetIndex()
        {
            if (position >= limit)
                throw new IndexOutOfRangeException();
            return position++;
        }

        protected int NextGetIndex(int nb)
        {
            if (limit - position < nb)
                throw new IndexOutOfRangeException();
            int p = position;
            position += nb;
            return p;
        }

        protected int NextPutIndex()
        {
            return NextGetIndex();
        }

        protected int NextPutIndex(int nb)
        {
            return NextGetIndex(nb);
        }

        protected int Index(int i)
        {
            return i + offset;
        }

        protected int CheckIndex(int i)
        {
            if(i < 0 || (i >= limit))
                throw new IndexOutOfRangeException();
            return i;
        }

        protected int CheckIndex(int i, int nb)
        {
            if ((i < 0) || (nb > limit - i))
                throw new IndexOutOfRangeException();
            return i;
        }

        public byte Read()
        {
            return buffer[Index(NextGetIndex())];
        }

        public byte this[int index]
        {
            get
            {
                return buffer[Index(CheckIndex(index))];
            }
            set
            {
             //   buffer[Index(NextPutIndex(index))] = value;
                if (index < 0 || index >= limit)
                    throw new IndexOutOfRangeException();
                buffer[Index(index)] = value;
            }
        }

        protected void CheckBounds(int off, int len, int size)
        {
            if ((off | len | (off + len) | (size - (off + len))) < 0)
                throw new IndexOutOfRangeException();
        }

        public ByteBuffer Read(byte[] buf)
        {
            return Read(buf, 0, buf.Length);
        }


        public ByteBuffer Read(byte[] dst, int offset, int length)
        {
            CheckBounds(offset, length, dst.Length);
            if (length > Remain)
                throw new ArgumentException();

            //for (int i = offset; i < offset + length; i++)
            //    dst[i] = Read();
            System.Array.Copy(buffer, Index(this.position), dst, offset, length);
            this.Position = Position + length;
            return this;
        }

        public ByteBuffer Write(byte x)
        {
            buffer[Index(NextPutIndex())] = x;
            return this;
        }

        public ByteBuffer Write(byte[] src, int offset, int length)
        {
            CheckBounds(offset, length, src.Length);

            if (length > Remain)
                throw new IndexOutOfRangeException();

            System.Array.Copy(src, offset, buffer, Index(this.Position), length);

            this.Position = this.Position + length;

            return this;
        }

        public ByteBuffer Write(ByteBuffer src)
        {
            if (src == this)
                throw new ArgumentException();
            int n = src.Remain;

            if (n > this.Remain)
                throw new IndexOutOfRangeException();

            for (int i = 0; i < n; i++)
                Write(src.Read());
            return this;
        }


        public ByteBuffer Write(byte[] data)
        {
            return Write(data, 0, data.Length);
        }

        public int Remain
        {
            get { return limit - position; }
        }
        
        public byte[] Buffer
        {
            get { return buffer; }
        }

        public int Offset
        {
            get { return offset; }
        }

        public bool HasBuffer
        {
            get { return true; }
        }

        public int Length
        {
            get { return capacity; }
        }
    }
}
/*

    private byte[] buffer;
        private int start;
        private int position;
       // private int capacity;
        private int length;

        public ByteBuffer(byte[] buffer, int start, int count)
        {
            this.buffer = buffer;
            this.start = position = start;
            //length = capacity = start + count;
            length = start + count;
        }

        public ByteBuffer(int capacity)
        {
            this.buffer = new byte[capacity];
            start = position = 0;
            length = capacity;
        }

        public ByteBuffer(byte[] buffer) : this(buffer, 0, buffer.Length)
        {

        }
        public void Write(byte[] data)
        {
            Write(data, 0, data.Length);
        }

        public void Write(byte[] data, int off, int count)
        {
            int len = data.Length;

            if ((off + count) > len || off < 0 || count < 0)
                throw new IndexOutOfRangeException();

            if (count > Remain)
                throw new Exception();

            for (int i = off; i < off + count; i++)
                buffer[position++] = data[i];
        }

        public ByteBuffer Write(int index, byte val)
        {
            if (index < 0 || index >= length)
                throw new IndexOutOfRangeException();
            buffer[index] = val;

            return this;
        }

        public int Remain
        {
            get { return length - position; }
        }

        public ByteBuffer Slice()  //This is not correct! Leon
        {
            ByteBuffer slice = new ByteBuffer(buffer, start + position, Remain);
            return slice;
        }

        public ByteBuffer Read(byte[] buf, int off, int count)
        {
            int len = buf.Length;

            if(off < 0 || count < 0 || (off + count > len))
                throw new IndexOutOfRangeException();
            if (count > Remain)
                throw new Exception();

            for (int i = off; i < off + count; i++)
                buf[i] = Read();
            return this;
        }

        public ByteBuffer Read(byte[] buf)
        {
            return Read(buf, 0, buf.Length);
        }

        public int Position
        {
            get { return position; }
            set
            {
                if (value < 0 || value > length)
                    throw new ArgumentException();
                position = value;
            }
        }

        public byte Read()
        {
            if (position == length)
                throw new Exception();

            return buffer[start + position++];
        }

        public byte[] Buffer
        {
            get { return buffer; }
        }

        public int Offset
        {
            get { return start; }
        }

        public bool HasBuffer
        {
            get { return true; }
        }

        public int Length
        {
            get { return length; }
        }

        //public int Position
        //{
        //    get { return position; }
        //}

        public byte this[int index]
        {
            get 
            {
                if (index < 0 || index > length)
                    throw new IndexOutOfRangeException();
                return buffer[index];
            }
        }

        public static ByteBuffer CreateBuffer(int capacity)
        {
            return new ByteBuffer(capacity);
        }

        public static ByteBuffer CreateBuffer(byte[] buffer, int start, int count)
        {
            int len = buffer.Length;

            if (start < 0 || count < 0 || (start + count > len))
                throw new IndexOutOfRangeException();
            ByteBuffer buf = new ByteBuffer(buffer);
            buf.position = start;
            buf.length = start + count;

            return buf;

        }

*/