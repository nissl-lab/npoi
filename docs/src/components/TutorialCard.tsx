import React from 'react';
import { motion } from 'motion/react';
import { Play, FileSpreadsheet, FileText, Settings, ExternalLink } from 'lucide-react';
import { Tutorial } from '../types';

interface TutorialCardProps {
  tutorial: Tutorial;
}

export const TutorialCard: React.FC<TutorialCardProps> = ({ tutorial }) => {
  const Icon = tutorial.category === 'Excel' 
    ? FileSpreadsheet 
    : tutorial.category === 'Word' 
      ? FileText 
      : Settings;

  const difficultyColor = {
    Beginner: 'bg-emerald-50 text-emerald-700 border-emerald-100',
    Intermediate: 'bg-amber-50 text-amber-700 border-amber-100',
    Advanced: 'bg-rose-50 text-rose-700 border-rose-100'
  }[tutorial.difficulty];

  return (
    <motion.div
      layout
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, scale: 0.95 }}
      whileHover={{ y: -4 }}
      className="group relative flex flex-col overflow-hidden rounded-2xl border border-zinc-200 bg-white p-6 shadow-sm transition-all hover:shadow-md"
    >
      <div className="mb-4 flex items-start justify-between">
        <div className={`flex h-12 w-12 items-center justify-center rounded-xl ${
          tutorial.category === 'Excel' ? 'bg-emerald-50 text-emerald-600' :
          tutorial.category === 'Word' ? 'bg-blue-50 text-blue-600' :
          'bg-zinc-50 text-zinc-600'
        }`}>
          <Icon size={24} />
        </div>
        <span className={`rounded-full border px-2.5 py-0.5 text-xs font-medium ${difficultyColor}`}>
          {tutorial.difficulty}
        </span>
      </div>

      <h3 className="mb-2 text-lg font-semibold leading-tight text-zinc-900 group-hover:text-zinc-900">
        {tutorial.title}
      </h3>
      
      <p className="mb-6 flex-grow text-sm leading-relaxed text-zinc-500">
        {tutorial.description}
      </p>

      <div className="flex items-center justify-between pt-4 border-t border-zinc-100">
        <span className="text-xs font-medium uppercase tracking-wider text-zinc-400">
          {tutorial.category}
        </span>
        <a
          href={tutorial.videoUrl}
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center gap-2 rounded-lg bg-zinc-900 px-4 py-2 text-sm font-medium text-white transition-all hover:bg-zinc-800 active:scale-95"
        >
          <Play size={14} fill="currentColor" />
          Watch Tutorial
          <ExternalLink size={12} className="opacity-50" />
        </a>
      </div>
    </motion.div>
  );
};
