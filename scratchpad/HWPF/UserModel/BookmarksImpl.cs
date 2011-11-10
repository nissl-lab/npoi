/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace NPOI.HWPF.usermodel;

using java.util.ArrayList;
using java.util.Arrays;
using java.util.Collections;
using java.util.HashMap;
using java.util.LinkedHashMap;
using java.util.LinkedList;
using java.util.List;
using java.util.Map;

using NPOI.HWPF.model.BookmarksTables;
using NPOI.HWPF.model.GenericPropertyNode;
using NPOI.HWPF.model.PropertyNode;

/**
 * Implementation of user-friendly interface for document bookmarks
 * 
 * @author Sergey Vladimirov (vlsergey {at} gmail {doc} com)
 */
public class BookmarksImpl : Bookmarks
{

    private BookmarksTables bookmarksTables;

    private Dictionary<Integer, List<GenericPropertyNode>> sortedDescriptors = null;

    private int[] sortedStartPositions = null;

    public BookmarksImpl( BookmarksTables bookmarksTables )
    {
        this.bookmarksTables = bookmarksTables;
    }

    private Bookmark GetBookmark( GenericPropertyNode first )
    {
        return new BookmarkImpl( first );
    }

    public Bookmark GetBookmark( int index )
    {
        GenericPropertyNode first = bookmarksTables
                .GetDescriptorFirst( index );
        return GetBookmark( first );
    }

    public List<Bookmark> GetBookmarksAt( int startCp )
    {
        updateSortedDescriptors();

        List<GenericPropertyNode> nodes = sortedDescriptors.Get( Integer
                .ValueOf( startCp ) );
        if ( nodes == null || nodes.isEmpty() )
            return Collections.emptyList();

        List<Bookmark> result = new ArrayList<Bookmark>( nodes.Count );
        for ( GenericPropertyNode node : nodes )
        {
            result.Add( GetBookmark( node ) );
        }
        return Collections.unmodifiableList( result );
    }

    public int GetBookmarksCount()
    {
        return bookmarksTables.GetDescriptorsFirstCount();
    }

    public Dictionary<Integer, List<Bookmark>> GetBookmarksStartedBetween(
            int startInclusive, int endExclusive )
    {
        updateSortedDescriptors();

        int startLookupIndex = Arrays.binarySearch( this.sortedStartPositions,
                startInclusive );
        if ( startLookupIndex < 0 )
            startLookupIndex = -( startLookupIndex + 1 );
        int endLookupIndex = Arrays.binarySearch( this.sortedStartPositions,
                endExclusive );
        if ( endLookupIndex < 0 )
            endLookupIndex = -( endLookupIndex + 1 );

        Dictionary<Integer, List<Bookmark>> result = new LinkedDictionary<Integer, List<Bookmark>>();
        for ( int LookupIndex = startLookupIndex; LookupIndex < endLookupIndex; LookupIndex++ )
        {
            int s = sortedStartPositions[LookupIndex];
            if ( s < startInclusive )
                continue;
            if ( s >= endExclusive )
                break;

            List<Bookmark> startedAt = GetBookmarksAt( s );
            if ( startedAt != null )
                result.Put( Int32.ValueOf( s ), startedAt );
        }

        return Collections.unmodifiableMap( result );
    }

    private void updateSortedDescriptors()
    {
        if ( sortedDescriptors != null )
            return;

        Dictionary<Integer, List<GenericPropertyNode>> result = new Dictionary<Integer, List<GenericPropertyNode>>();
        for ( int b = 0; b < bookmarksTables.GetDescriptorsFirstCount(); b++ )
        {
            GenericPropertyNode property = bookmarksTables
                    .GetDescriptorFirst( b );
            Integer positionKey = Int32.ValueOf( property.Start );
            List<GenericPropertyNode> atPositionList = result.Get( positionKey );
            if ( atPositionList == null )
            {
                atPositionList = new LinkedList<GenericPropertyNode>();
                result.Put( positionKey, atPositionList );
            }
            atPositionList.Add( property );
        }

        int counter = 0;
        int[] indices = new int[result.Count];
        for ( Map.Entry<Integer, List<GenericPropertyNode>> entry : result
                .entrySet() )
        {
            indices[counter++] = entry.GetKey().intValue();
            List<GenericPropertyNode> updated = new ArrayList<GenericPropertyNode>(
                    entry.GetValue() );
            Collections.sort( updated, PropertyNode.EndComparator.instance );
            entry.SetValue( updated );
        }
        Arrays.sort( indices );

        this.sortedDescriptors = result;
        this.sortedStartPositions = indices;
    }

    private class BookmarkImpl : Bookmark
    {
        private GenericPropertyNode first;

        private BookmarkImpl( GenericPropertyNode first )
        {
            this.first = first;
        }

        @Override
        public bool Equals( Object obj )
        {
            if ( this == obj )
                return true;
            if ( obj == null )
                return false;
            if ( GetClass() != obj.GetClass() )
                return false;
            BookmarkImpl other = (BookmarkImpl) obj;
            if ( first == null )
            {
                if ( other.first != null )
                    return false;
            }
            else if ( !first.Equals( other.first ) )
                return false;
            return true;
        }

        public int GetEnd()
        {
            int currentIndex = bookmarksTables.GetDescriptorFirstIndex( first );
            try
            {
                GenericPropertyNode descriptorLim = bookmarksTables
                        .GetDescriptorLim( currentIndex );
                return descriptorLim.Start;
            }
            catch ( IndexOutOfBoundsException exc )
            {
                return first.End;
            }
        }

        public String GetName()
        {
            int currentIndex = bookmarksTables.GetDescriptorFirstIndex( first );
            try
            {
                return bookmarksTables.GetName( currentIndex );
            }
            catch ( ArrayIndexOutOfBoundsException exc )
            {
                return "";
            }
        }

        public int GetStart()
        {
            return first.Start;
        }

        @Override
        public int hashCode()
        {
            return 31 + ( first == null ? 0 : first.HashCode() );
        }

        public void SetName( String name )
        {
            int currentIndex = bookmarksTables.GetDescriptorFirstIndex( first );
            bookmarksTables.SetName( currentIndex, name );
        }

        @Override
        public String ToString()
        {
            return "Bookmark [" + GetStart() + "; " + GetEnd() + "): name: "
                    + GetName();
        }

    }
}


