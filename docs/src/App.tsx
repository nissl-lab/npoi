import { useState, useMemo } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Search, Filter, LayoutGrid, List as ListIcon, Sparkles } from 'lucide-react';
import { Navbar } from './components/Navbar';
import { TutorialCard } from './components/TutorialCard';
import { TUTORIALS } from './constants';
import { Tutorial } from './types';

export default function App() {
  const [searchQuery, setSearchQuery] = useState('');
  const [activeCategory, setActiveCategory] = useState<string>('All');
  const [viewMode, setViewMode] = useState<'grid' | 'list'>('grid');

  const categories = ['All', 'Excel', 'Word', 'General'];

  const filteredTutorials = useMemo(() => {
    return TUTORIALS.filter((t) => {
      const matchesSearch = t.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
                          t.description.toLowerCase().includes(searchQuery.toLowerCase());
      const matchesCategory = activeCategory === 'All' || t.category === activeCategory;
      return matchesSearch && matchesCategory;
    });
  }, [searchQuery, activeCategory]);

  return (
    <div className="min-h-screen bg-zinc-50 font-sans text-zinc-900">
      <Navbar />

      <main className="mx-auto max-w-7xl px-4 py-12 sm:px-6 lg:px-8">
        {/* Hero Section */}
        <header className="mb-16 text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="inline-flex items-center gap-2 rounded-full bg-zinc-900 px-3 py-1 text-xs font-medium text-white mb-6"
          >
            <Sparkles size={14} className="text-amber-400" />
            <span>Master .NET Office Automation</span>
          </motion.div>
          <motion.h1
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 }}
            className="mb-6 text-5xl font-extrabold tracking-tight text-zinc-900 sm:text-6xl"
          >
            NPOI <span className="text-zinc-500 italic">Tutorials</span>
          </motion.h1>
          <motion.p
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.2 }}
            className="mx-auto max-w-2xl text-lg text-zinc-600"
          >
            A curated collection of step-by-step guides to help you master Microsoft Office file manipulation using the powerful NPOI library.
          </motion.p>
        </header>

        {/* Controls */}
        <div className="mb-12 flex flex-col gap-6 sm:flex-row sm:items-center sm:justify-between">
          <div className="relative flex-grow max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" size={18} />
            <input
              type="text"
              placeholder="Search tutorials..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full rounded-xl border border-zinc-200 bg-white py-2.5 pl-10 pr-4 text-sm outline-none transition-all focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900"
            />
          </div>

          <div className="flex items-center gap-4 overflow-x-auto pb-2 sm:pb-0">
            <div className="flex items-center gap-1 rounded-xl border border-zinc-200 bg-white p-1">
              {categories.map((cat) => (
                <button
                  key={cat}
                  onClick={() => setActiveCategory(cat)}
                  className={`rounded-lg px-4 py-1.5 text-sm font-medium transition-all ${
                    activeCategory === cat
                      ? 'bg-zinc-900 text-white shadow-sm'
                      : 'text-zinc-500 hover:bg-zinc-50 hover:text-zinc-900'
                  }`}
                >
                  {cat}
                </button>
              ))}
            </div>

            <div className="hidden h-8 w-px bg-zinc-200 sm:block" />

            <div className="flex items-center gap-1 rounded-xl border border-zinc-200 bg-white p-1">
              <button
                onClick={() => setViewMode('grid')}
                className={`rounded-lg p-1.5 transition-all ${
                  viewMode === 'grid' ? 'bg-zinc-100 text-zinc-900' : 'text-zinc-400 hover:text-zinc-600'
                }`}
              >
                <LayoutGrid size={18} />
              </button>
              <button
                onClick={() => setViewMode('list')}
                className={`rounded-lg p-1.5 transition-all ${
                  viewMode === 'list' ? 'bg-zinc-100 text-zinc-900' : 'text-zinc-400 hover:text-zinc-600'
                }`}
              >
                <ListIcon size={18} />
              </button>
            </div>
          </div>
        </div>

        {/* Tutorial Grid */}
        <motion.div
          layout
          className={
            viewMode === 'grid'
              ? "grid gap-6 sm:grid-cols-2 lg:grid-cols-3"
              : "flex flex-col gap-4"
          }
        >
          <AnimatePresence mode="popLayout">
            {filteredTutorials.map((tutorial) => (
              <TutorialCard key={tutorial.id} tutorial={tutorial} />
            ))}
          </AnimatePresence>
        </motion.div>

        {filteredTutorials.length === 0 && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            className="flex flex-col items-center justify-center py-20 text-center"
          >
            <div className="mb-4 rounded-full bg-zinc-100 p-4 text-zinc-400">
              <Search size={32} />
            </div>
            <h3 className="text-lg font-semibold text-zinc-900">No tutorials found</h3>
            <p className="text-zinc-500">Try adjusting your search or filter to find what you're looking for.</p>
            <button
              onClick={() => {
                setSearchQuery('');
                setActiveCategory('All');
              }}
              className="mt-6 text-sm font-medium text-zinc-900 underline underline-offset-4"
            >
              Clear all filters
            </button>
          </motion.div>
        )}
      </main>

      {/* Footer */}
      <footer className="mt-20 border-t border-zinc-200 bg-white py-12">
        <div className="mx-auto max-w-7xl px-4 text-center sm:px-6 lg:px-8">
          <p className="text-sm text-zinc-500">
            Built for the .NET community. Based on tutorials by <a href="https://github.com/nissl-lab" className="font-medium text-zinc-900 hover:underline">nissl-lab</a>.
          </p>
          <div className="mt-4 flex justify-center gap-6">
            <a href="#" className="text-zinc-400 hover:text-zinc-600 transition-colors">Documentation</a>
            <a href="#" className="text-zinc-400 hover:text-zinc-600 transition-colors">NuGet Package</a>
            <a href="#" className="text-zinc-400 hover:text-zinc-600 transition-colors">Support</a>
          </div>
        </div>
      </footer>
    </div>
  );
}

