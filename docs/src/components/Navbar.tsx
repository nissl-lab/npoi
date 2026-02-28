import React from 'react';
import { BookOpen, Github, Youtube } from 'lucide-react';

export const Navbar: React.FC = () => {
  return (
    <nav className="sticky top-0 z-50 w-full border-b border-zinc-200 bg-white/80 backdrop-blur-md">
      <div className="mx-auto flex h-16 max-w-7xl items-center justify-between px-4 sm:px-6 lg:px-8">
        <div className="flex items-center gap-2">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-zinc-900 text-white">
            <BookOpen size={24} />
          </div>
          <span className="text-xl font-bold tracking-tight text-zinc-900">NPOI Mastery</span>
        </div>
        
        <div className="flex items-center gap-4">
          <a
            href="https://github.com/nissl-lab/npoi-tutorial"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-2 rounded-lg px-3 py-2 text-sm font-medium text-zinc-600 transition-colors hover:bg-zinc-100 hover:text-zinc-900"
          >
            <Github size={18} />
            <span className="hidden sm:inline">GitHub</span>
          </a>
          <a
            href="https://www.youtube.com/playlist?list=PL7J6yRMWV1ot32hsgCZJ5sI4Fp2QKovCM"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-2 rounded-lg bg-red-50 px-3 py-2 text-sm font-medium text-red-600 transition-colors hover:bg-red-100"
          >
            <Youtube size={18} />
            <span className="hidden sm:inline">Playlist</span>
          </a>
        </div>
      </div>
    </nav>
  );
};
