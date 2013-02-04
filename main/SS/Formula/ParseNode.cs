/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Function;
    /**
     * Represents a syntactic element from a formula by encapsulating the corresponding <c>Ptg</c>
     * Token.  Each <c>ParseNode</c> may have child <c>ParseNode</c>s in the case when the wrapped
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
                tokenCount += children[i].TokenCount;
            }
            if (_isIf)
            {
                // there will be 2 or 3 extra tAttr Tokens according To whether the false param is present
                tokenCount += children.Length;
            }
            _tokenCount = tokenCount;
        }
        public ParseNode(Ptg token)
            : this(token, EMPTY_ARRAY)
        {

        }
        public ParseNode(Ptg token, ParseNode child0)
            : this(token, new ParseNode[] { child0, })
        {

        }
        public ParseNode(Ptg token, ParseNode child0, ParseNode child1)
            : this(token, new ParseNode[] { child0, child1, })
        {

        }
        private int TokenCount
        {
            get
            {
                return _tokenCount;
            }
        }
        public int EncodedSize
        {
            get
            {
                int result = _token is ArrayPtg ? ArrayPtg.PLAIN_TOKEN_SIZE : _token.Size;
                for (int i = 0; i < _children.Length; i++)
                {
                    result += _children[i].EncodedSize;
                }
                return result;
            }
        }

        /**
         * Collects the array of <c>Ptg</c> Tokens for the specified tree.
         */
        public static Ptg[] ToTokenArray(ParseNode rootNode)
        {
            TokenCollector temp = new TokenCollector(rootNode.TokenCount);
            rootNode.CollectPtgs(temp);
            return temp.GetResult();
        }
        private void CollectPtgs(TokenCollector temp)
        {
            if (IsIf(_token))
            {
                CollectIfPtgs(temp);
                return;
            }
            bool isPreFixOperator = _token is MemFuncPtg || _token is MemAreaPtg;
		    if (isPreFixOperator) {
			    temp.Add(_token);
		    }
            for (int i = 0; i < GetChildren().Length; i++)
            {
                GetChildren()[i].CollectPtgs(temp);
            }
            if(!isPreFixOperator)
            {
                temp.Add(_token);
            }
        }
        /**
         * The IF() function Gets marked up with two or three tAttr Tokens.
         * Similar logic will be required for CHOOSE() when it is supported
         * 
         * See excelfileformat.pdf sec 3.10.5 "tAttr (19H)
         */
        private void CollectIfPtgs(TokenCollector temp)
        {

            // condition goes first
            GetChildren()[0].CollectPtgs(temp);

            // placeholder for tAttrIf
            int ifAttrIndex = temp.CreatePlaceholder();

            // true parameter
            GetChildren()[1].CollectPtgs(temp);

            // placeholder for first skip attr
            int skipAfterTrueParamIndex = temp.CreatePlaceholder();
            int trueParamSize = temp.sumTokenSizes(ifAttrIndex + 1, skipAfterTrueParamIndex);

            AttrPtg attrIf = AttrPtg.CreateIf(trueParamSize + 4);// distance to start of false parameter/tFuncVar. +4 for tAttrSkip after true

            if (GetChildren().Length > 2)
            {
                // false param present

                // false parameter
                GetChildren()[2].CollectPtgs(temp);

                int skipAfterFalseParamIndex = temp.CreatePlaceholder();
                int falseParamSize = temp.sumTokenSizes(skipAfterTrueParamIndex + 1, skipAfterFalseParamIndex);

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

            temp.Add(GetToken());
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

        public Ptg GetToken()
        {
            return _token;
        }

        public ParseNode[] GetChildren()
        {
            return _children;
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

            public int sumTokenSizes(int fromIx, int ToIx)
            {
                int result = 0;
                for (int i = fromIx; i < ToIx; i++)
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

            public Ptg[] GetResult()
            {
                return _ptgs;
            }
        }
    }
}