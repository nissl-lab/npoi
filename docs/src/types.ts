export interface Tutorial {
  id: number;
  title: string;
  description: string;
  videoUrl: string;
  category: 'Excel' | 'Word' | 'General';
  difficulty: 'Beginner' | 'Intermediate' | 'Advanced';
}
