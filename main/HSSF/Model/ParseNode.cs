/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Model
{
    using System;

    using NPOI.SS.Formula.Function;
    using NPOI.SS.Formula.PTG;
    /**
     * Represents a syntactic element from a formula by encapsulating the corresponding <c>Ptg</c>
     * token.  Each <c>ParseNode</c> may have child <c>ParseNode</c>s in the case when the wrapped
     * <c>Ptg</c> is non-atomic.
     * 
     * @author Josh Micich
     */
    class ParseNode
    {

        public static ParseNode[] EMPTY_ARRAY = { };
        private Ptg _token;
        private ParseNode[] _children;
        private bool _isIf;
        private int _tokenCount;

        public ParseNode(Ptg token, ParseNode[] children)
        {
            _token = token;
            _children = children;
            _isIf = IsIf(token);
            int tokenCount = 1;
            for (int i = 0; i < children.Length; i++)
            {
                tokenCount += children[i].GetTokenCount();
            }
            if (_isIf)
            {
                // there will be 2 or 3 extra tAttr tokens according to whether the false param is present
                tokenCount += children.Length;
            }
            _tokenCount = tokenCount;
        }
        public ParseNode(Ptg token):this(token, EMPTY_ARRAY) 
        {
            
        }
        public ParseNode(Ptg token, ParseNode child0):this(token, new ParseNode[] { child0, })
        {
            
        }
        public ParseNode(Ptg token, ParseNode child0, ParseNode child1): this(token, new ParseNode[] { child0, child1, })
        {
           
        }
        private int GetTokenCount()
        {
            return _tokenCount;
        }

        /// <summary>
        /// Collects the array of Ptg
        ///  tokens for the specified tree.
        /// </summary>
        /// <param name="rootNode">The root node.</param>
        /// <returns></returns>
        public static Ptg[] ToTokenArray(ParseNode rootNode)
        {
            TokenCollector temp = new TokenCollector(rootNode.GetTokenCount());
            rootNode.CollectPtgs(temp);
            return temp.Result;
        }
        private void CollectPtgs(TokenCollector temp)
        {
            if (IsIf(this.Token))
            {
                CollectIfPtgs(temp);
                return;
            }
            for (int i = 0; i < Children.Length; i++)
            {
                Children[i].CollectPtgs(temp);
            }
            temp.Add(this.Token);
        }
        /// <summary>
        /// The IF() function Gets marked up with two or three tAttr tokens.
        /// Similar logic will be required for CHOOSE() when it is supported
        /// See excelfileformat.pdf sec 3.10.5 "tAttr (19H)
        /// </summary>
        /// <param name="temp">The temp.</param>
        private void CollectIfPtgs(TokenCollector temp)
        {

            // condition goes first
            Children[0].CollectPtgs(temp);

            // placeholder for tAttrIf
            int ifAttrIndex = temp.CreatePlaceholder();

            // true parameter
            Children[1].CollectPtgs(temp);

            // placeholder for first skip attr
            int skipAfterTrueParamIndex = temp.CreatePlaceholder();
            int trueParamSize = temp.SumTokenSizes(ifAttrIndex + 1, skipAfterTrueParamIndex);

            AttrPtg attrIf = AttrPtg.CreateIf(trueParamSize + 4); // distance to start of false parameter/tFuncVar. +4 for tAttrSkip after true
            
            if (Children.Length > 2)
            {
                // false param present

                // false parameter
                Children[2].CollectPtgs(temp);

                int skipAfterFalseParamIndex = temp.CreatePlaceholder();
                int falseParamSize = temp.SumTokenSizes(skipAfterTrueParamIndex + 1, skipAfterFalseParamIndex);

                AttrPtg attrSkipAfterTrue = AttrPtg.CreateSkip(falseParamSize + 4 + 4 - 1); // 1 less than distance to end of if FuncVar(size=4). +4 for attr skip before
                AttrPtg attrSkipAfterFalse = AttrPtg.CreateSkip(4 - 1); // 1 less than distance to end of if FuncVar(size=4).

                temp.SetPlaceholder(ifAttrIndex, attrIf);
                temp.SetPlaceholder(skipAfterTrueParamIndex, attrSkipAfterTrue);
                temp.SetPlaceholder(skipAfterFalseParamIndex, attrSkipAfterFalse);
            }
            else
            {
                // false parameter not present
                AttrPtg attrSkipAfterTrue = AttrPtg.CreateSkip(4 - 1); // 1 less than distance to end of if FuncVar(size=4).

                temp.SetPlaceholder(ifAttrIndex, attrIf);
                temp.SetPlaceholder(skipAfterTrueParamIndex, attrSkipAfterTrue);
            }

            temp.Add(_token);
        }

        private static bool IsIf(Ptg token)
        {
            if (token is FuncVarPtg)
            {
                FuncVarPtg func = (FuncVarPtg)token;
                if (FunctionMetadataRegistry.FUNCTION_NAME_IF.Equals(func.Name))
                {
                    return true;
                }
            }
            return false;
        }

        public Ptg Token
        {
            get
            {
                return _token;
            }
        }

        public ParseNode[] Children
        {
            get { return _children; }
        }

        private class TokenCollector
        {

            private Ptg[] _ptgs;
            private int _offset;

            public TokenCollector(int tokenCount)
            {
                _ptgs = new Ptg[tokenCount];
                _offset = 0;
            }

            public int SumTokenSizes(int fromIx, int toIx)
            {
                int result = 0;
                for (int i = fromIx; i < toIx; i++)
                {
                    result += _ptgs[i].Size;
                }
                return result;
            }

            public int CreatePlaceholder()
            {
                return _offset++;
            }

            public void Add(Ptg token)
            {
                if (token == null)
                {
                    throw new ArgumentException("token must not be null");
                }
                _ptgs[_offset] = token;
                _offset++;
            }

            public void SetPlaceholder(int index, Ptg token)
            {
                if (_ptgs[index] != null)
                {
                    throw new InvalidOperationException("Invalid placeholder index (" + index + ")");
                }
                _ptgs[index] = token;
            }

            public Ptg[] Result
            {
                get
                {
                    return _ptgs;
                }
            }
        }
    }
}