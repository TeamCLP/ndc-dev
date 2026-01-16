// js/components/MusicPlayer.js
export class MusicPlayer {
    constructor() {
        // Vérifier si une instance existe déjà
        if (window.musicPlayerInstance) {
            console.warn('⚠️ MusicPlayer déjà initialisé, réutilisation de l\'instance existante');
            return window.musicPlayerInstance;
        }
        
        this.isPlaying = false;
        this.currentTrack = 0;
        this.audio = new Audio();
        this.isMinimized = true;
        this.audioUnlocked = false;
        
        // Configuration de l'audio
        this.audio.preload = 'auto';
        this.audio.crossOrigin = 'anonymous';
        
        // Sauvegarder l'instance globalement
        window.musicPlayerInstance = this;
        
        // Playlist - Ajoutez vos MP3 ici
        this.playlist = [
            {
                title: "L'IA Architecte",
                artist: "Quentin",
                src: "music/L-IA-Architecte.mp3",
                duration: "2:32"
            },
            {
                title: "L'Architecte Invisible",
                artist: "Quentin",
                src: "music/L-Architecte-Invisible.mp3",
                duration: "2:19"
            },
            {
                title: "L'IA qui rêve",
                artist: "Quentin",
                src: "music/L-IA-qui-rêve.mp3",
                duration: "2:56"
            },
            {
                title: "L'Étincelle",
                artist: "Quentin",
                src: "music/L-Étincelle.mp3",
                duration: "2:29"
            }
        ];
        
        this.init();
    }
    
    init() {
        console.log('🎵 MusicPlayer.init() appelé');
        this.createPlayerHTML();
        console.log('  ✓ HTML créé');
        this.bindEvents();
        console.log('  ✓ Events liés');
        this.loadTrack(0);
        console.log('  ✓ Première piste chargée');
        
        // Vérifier que les éléments existent
        const button = document.getElementById('musicButton');
        const player = document.getElementById('musicPlayer');
        console.log('  - Bouton musique:', button ? 'TROUVÉ' : 'INTROUVABLE');
        console.log('  - Lecteur:', player ? 'TROUVÉ' : 'INTROUVABLE');
        
        // Restaurer les préférences
        const savedVolume = localStorage.getItem('scribe_music_volume');
        if (savedVolume) {
            this.audio.volume = parseFloat(savedVolume);
            this.updateVolumeUI();
        } else {
            // Volume par défaut à 50%
            this.audio.volume = 0.5;
            this.updateVolumeUI();
        }
        
        const savedTrack = localStorage.getItem('scribe_music_track');
        if (savedTrack) {
            this.currentTrack = parseInt(savedTrack);
            this.loadTrack(this.currentTrack);
        }
        
        // Ajouter un indicateur de statut audio
        console.log('🎵 MusicPlayer initialisé');
        console.log('🔊 Volume initial:', Math.round(this.audio.volume * 100) + '%');
        console.log('🎵 Piste chargée:', this.playlist[this.currentTrack].title);
    }
    
    createPlayerHTML() {
        // Vérifier si les éléments existent déjà
        if (document.getElementById('musicButton')) {
            console.log('⚠️ Elements MusicPlayer déjà présents, nettoyage...');
            document.getElementById('musicButton')?.remove();
            document.getElementById('musicPlayer')?.remove();
        }
        
        // Bouton flottant pour ouvrir le lecteur
        const musicButton = document.createElement('button');
        musicButton.className = 'music-button';
        musicButton.id = 'musicButton';
        musicButton.innerHTML = '🎵';
        musicButton.title = 'Lecteur de musique';
        document.body.appendChild(musicButton);
        
        // Lecteur de musique
        const playerContainer = document.createElement('div');
        playerContainer.className = 'music-player minimized';
        playerContainer.id = 'musicPlayer';
        playerContainer.innerHTML = `
            <div class="music-player-header">
                <h4>🎵 Musique d'ambiance</h4>
                <button class="music-close" id="musicClose">×</button>
            </div>
            
            <div class="music-player-content">
                <div class="track-info">
                    <div class="track-title" id="trackTitle">Chargement...</div>
                    <div class="track-artist" id="trackArtist">...</div>
                </div>
                
                <div class="progress-container">
                    <span class="time-current" id="timeCurrent">0:00</span>
                    <div class="progress-bar" id="progressBar">
                        <div class="progress-fill" id="progressFill"></div>
                    </div>
                    <span class="time-total" id="timeTotal">0:00</span>
                </div>
                
                <div class="music-controls">
                    <button class="control-btn" id="prevBtn" title="Précédent">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                            <path d="M6 6h2v12H6zm3.5 6l8.5 6V6z"/>
                        </svg>
                    </button>
                    <button class="control-btn play-btn" id="playBtn" title="Lecture/Pause">
                        <svg class="play-icon" width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
                            <path d="M8 5v14l11-7z"/>
                        </svg>
                        <svg class="pause-icon" width="24" height="24" viewBox="0 0 24 24" fill="currentColor" style="display: none;">
                            <path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/>
                        </svg>
                    </button>
                    <button class="control-btn" id="nextBtn" title="Suivant">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                            <path d="M6 18l8.5-6L6 6v12zM16 6v12h2V6h-2z"/>
                        </svg>
                    </button>
                </div>
                
                <div class="volume-container">
                    <button class="volume-btn" id="volumeBtn">
                        <svg class="volume-icon" width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                            <path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3c0-1.77-1.02-3.29-2.5-4.03v8.05c1.48-.73 2.5-2.25 2.5-4.02zM14 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77z"/>
                        </svg>
                        <svg class="mute-icon" width="20" height="20" viewBox="0 0 24 24" fill="currentColor" style="display: none;">
                            <path d="M16.5 12c0-1.77-1.02-3.29-2.5-4.03v2.21l2.45 2.45c.03-.2.05-.41.05-.63zm2.5 0c0 .94-.2 1.82-.54 2.64l1.51 1.51C20.63 14.91 21 13.5 21 12c0-4.28-2.99-7.86-7-8.77v2.06c2.89.86 5 3.54 5 6.71zM4.27 3L3 4.27 7.73 9H3v6h4l5 5v-6.73l4.25 4.25c-.67.52-1.42.93-2.25 1.18v2.06c1.38-.31 2.63-.95 3.69-1.81L19.73 21 21 19.73l-9-9L4.27 3zM12 4L9.91 6.09 12 8.18V4z"/>
                        </svg>
                    </button>
                    <input type="range" class="volume-slider" id="volumeSlider" min="0" max="100" value="50">
                </div>
                
                <div class="playlist-container">
                    <h5>Playlist</h5>
                    <div class="playlist" id="playlist"></div>
                </div>
            </div>
        `;
        
        document.body.appendChild(playerContainer);
        
        // Ajouter les styles CSS
        this.addStyles();
        
        // Remplir la playlist
        this.updatePlaylistUI();
    }
    
    addStyles() {
        const style = document.createElement('style');
        style.textContent = `
            /* Bouton flottant */
            .music-button {
                position: fixed;
                bottom: 80px;
                right: 20px;
                width: 50px;
                height: 50px;
                border-radius: 50%;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                border: none;
                color: white;
                font-size: 24px;
                cursor: pointer;
                box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
                transition: all 0.3s ease;
                z-index: 1000;
                animation: pulse 2s infinite;
            }
            
            @keyframes pulse {
                0% { transform: scale(1); }
                50% { transform: scale(1.05); }
                100% { transform: scale(1); }
            }
            
            .music-button:hover {
                transform: scale(1.1);
                box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
            }
            
            .music-button.playing {
                animation: rotate 3s linear infinite, pulse 2s infinite;
            }
            
            @keyframes rotate {
                from { transform: rotate(0deg); }
                to { transform: rotate(360deg); }
            }
            
            /* Lecteur de musique */
            .music-player {
                position: fixed;
                bottom: 140px;
                right: 20px;
                width: 350px;
                background: rgba(255, 255, 255, 0.98);
                border-radius: 15px;
                box-shadow: 0 10px 40px rgba(0, 0, 0, 0.15);
                transition: all 0.3s ease;
                z-index: 999;
                backdrop-filter: blur(10px);
            }
            
            .music-player.minimized {
                transform: translateY(20px);
                opacity: 0;
                pointer-events: none;
            }
            
            .music-player-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 15px 20px;
                border-bottom: 1px solid #e5e7eb;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                border-radius: 15px 15px 0 0;
                color: white;
            }
            
            .music-player-header h4 {
                margin: 0;
                font-size: 16px;
                font-weight: 600;
            }
            
            .music-close {
                background: none;
                border: none;
                color: white;
                font-size: 24px;
                cursor: pointer;
                opacity: 0.8;
                transition: opacity 0.2s;
                padding: 0;
                width: 30px;
                height: 30px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 50%;
            }
            
            .music-close:hover {
                opacity: 1;
                background: rgba(255, 255, 255, 0.2);
            }
            
            .music-player-content {
                padding: 20px;
            }
            
            .track-info {
                text-align: center;
                margin-bottom: 20px;
            }
            
            .track-title {
                font-size: 18px;
                font-weight: 600;
                color: #1f2937;
                margin-bottom: 5px;
            }
            
            .track-artist {
                font-size: 14px;
                color: #6b7280;
            }
            
            .progress-container {
                display: flex;
                align-items: center;
                gap: 10px;
                margin-bottom: 25px;
            }
            
            .progress-bar {
                flex: 1;
                height: 5px;
                background: #e5e7eb;
                border-radius: 5px;
                cursor: pointer;
                position: relative;
                overflow: hidden;
            }
            
            .progress-fill {
                height: 100%;
                background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
                width: 0%;
                transition: width 0.1s;
                position: relative;
            }
            
            .progress-fill::after {
                content: '';
                position: absolute;
                right: -6px;
                top: 50%;
                transform: translateY(-50%);
                width: 12px;
                height: 12px;
                background: white;
                border: 3px solid #764ba2;
                border-radius: 50%;
                opacity: 0;
                transition: opacity 0.2s;
            }
            
            .progress-bar:hover .progress-fill::after {
                opacity: 1;
            }
            
            .time-current, .time-total {
                font-size: 12px;
                color: #6b7280;
                font-family: monospace;
                min-width: 35px;
            }
            
            .music-controls {
                display: flex;
                justify-content: center;
                align-items: center;
                gap: 15px;
                margin-bottom: 20px;
            }
            
            .control-btn {
                background: #f3f4f6;
                border: none;
                width: 45px;
                height: 45px;
                border-radius: 50%;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: all 0.2s;
                color: #4b5563;
            }
            
            .control-btn:hover {
                background: #e5e7eb;
                transform: scale(1.05);
            }
            
            .play-btn {
                width: 55px;
                height: 55px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
            }
            
            .play-btn:hover {
                transform: scale(1.1);
                box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
            }
            
            .volume-container {
                display: flex;
                align-items: center;
                gap: 10px;
                margin-bottom: 20px;
            }
            
            .volume-btn {
                background: none;
                border: none;
                color: #6b7280;
                cursor: pointer;
                padding: 5px;
                display: flex;
                align-items: center;
                justify-content: center;
            }
            
            .volume-btn:hover {
                color: #667eea;
            }
            
            .volume-slider {
                flex: 1;
                height: 5px;
                -webkit-appearance: none;
                appearance: none;
                background: #e5e7eb;
                border-radius: 5px;
                outline: none;
            }
            
            .volume-slider::-webkit-slider-thumb {
                -webkit-appearance: none;
                appearance: none;
                width: 15px;
                height: 15px;
                background: #667eea;
                border-radius: 50%;
                cursor: pointer;
                transition: all 0.2s;
            }
            
            .volume-slider::-webkit-slider-thumb:hover {
                background: #764ba2;
                transform: scale(1.2);
            }
            
            .playlist-container {
                margin-top: 25px;
                border-top: 1px solid #e5e7eb;
                padding-top: 15px;
            }
            
            .playlist-container h5 {
                margin: 0 0 10px 0;
                font-size: 14px;
                color: #6b7280;
                font-weight: 600;
            }
            
            .playlist {
                max-height: 150px;
                overflow-y: auto;
            }
            
            .playlist-item {
                padding: 10px;
                cursor: pointer;
                border-radius: 8px;
                transition: all 0.2s;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }
            
            .playlist-item:hover {
                background: #f3f4f6;
            }
            
            .playlist-item.active {
                background: linear-gradient(135deg, rgba(102, 126, 234, 0.1) 0%, rgba(118, 75, 162, 0.1) 100%);
                color: #667eea;
            }
            
            .playlist-item-title {
                font-size: 14px;
                font-weight: 500;
            }
            
            .playlist-item-duration {
                font-size: 12px;
                color: #9ca3af;
            }
            
            /* Dark mode support */
            @media (prefers-color-scheme: dark) {
                .music-player {
                    background: rgba(31, 41, 55, 0.98);
                    color: #f3f4f6;
                }
                
                .track-title {
                    color: #f3f4f6;
                }
                
                .track-artist {
                    color: #9ca3af;
                }
                
                .progress-bar {
                    background: #374151;
                }
                
                .control-btn {
                    background: #374151;
                    color: #d1d5db;
                }
                
                .control-btn:hover {
                    background: #4b5563;
                }
                
                .volume-slider {
                    background: #374151;
                }
                
                .playlist-item:hover {
                    background: #374151;
                }
            }
        `;
        document.head.appendChild(style);
    }
    
    bindEvents() {
        console.log('🎵 MusicPlayer.bindEvents() appelé');
        
        // Bouton musique
        const musicButton = document.getElementById('musicButton');
        if (musicButton) {
            console.log('  ✓ Bouton musique trouvé, ajout du listener');
            musicButton.addEventListener('click', () => {
                console.log('🎵 CLIC sur bouton musique!');
                this.togglePlayer();
            });
        } else {
            console.error('  ❌ Bouton musique introuvable!');
        }
        
        // Fermer le lecteur
        const closeButton = document.getElementById('musicClose');
        if (closeButton) {
            closeButton.addEventListener('click', () => {
                console.log('🎵 Fermeture du lecteur');
                this.togglePlayer();
            });
        }
        
        // Contrôles
        document.getElementById('playBtn').addEventListener('click', () => {
            this.togglePlay();
        });
        
        document.getElementById('prevBtn').addEventListener('click', () => {
            this.previousTrack();
        });
        
        document.getElementById('nextBtn').addEventListener('click', () => {
            this.nextTrack();
        });
        
        // Progress bar
        document.getElementById('progressBar').addEventListener('click', (e) => {
            this.seekTo(e);
        });
        
        // Volume
        document.getElementById('volumeBtn').addEventListener('click', () => {
            this.toggleMute();
        });
        
        document.getElementById('volumeSlider').addEventListener('input', (e) => {
            this.setVolume(e.target.value / 100);
        });
        
        // Audio events
        this.audio.addEventListener('timeupdate', () => {
            this.updateProgress();
        });
        
        this.audio.addEventListener('ended', () => {
            this.nextTrack();
        });
        
        this.audio.addEventListener('loadedmetadata', () => {
            this.updateDuration();
        });
        
        // Gestion des erreurs de chargement
        this.audio.addEventListener('error', (e) => {
            console.error('❌ Erreur de chargement audio:', e);
            console.error('Source tentée:', this.audio.src);
            alert(`Erreur de chargement du fichier audio: ${this.playlist[this.currentTrack].title}`);
        });
        
        // Playlist clicks
        document.getElementById('playlist').addEventListener('click', (e) => {
            const item = e.target.closest('.playlist-item');
            if (item) {
                const index = parseInt(item.dataset.index);
                this.loadTrack(index);
                // Toujours essayer de jouer après un clic utilisateur
                setTimeout(() => {
                    this.play();
                }, 100);
            }
        });
        
        // Premier clic pour débloquer l'audio sur mobile/certains navigateurs
        document.body.addEventListener('click', () => {
            if (!this.audioUnlocked) {
                this.unlockAudio();
                this.audioUnlocked = true;
            }
        }, { once: true });
    }
    
    togglePlayer() {
        const player = document.getElementById('musicPlayer');
        this.isMinimized = !this.isMinimized;
        
        if (this.isMinimized) {
            player.classList.add('minimized');
        } else {
            player.classList.remove('minimized');
            // Test automatique de l'audio à l'ouverture
            console.log('🎵 Lecteur ouvert - État audio:');
            console.log('  - Source:', this.audio.src);
            console.log('  - Volume:', this.audio.volume);
            console.log('  - Muted:', this.audio.muted);
            console.log('  - Ready state:', this.audio.readyState);
            console.log('  - Network state:', this.audio.networkState);
        }
    }
    
    togglePlay() {
        if (this.isPlaying) {
            this.pause();
        } else {
            this.play();
        }
    }
    
    play() {
        // Vérifier l'état avant de jouer
        console.log('🎵 Tentative de lecture...');
        console.log('  - Fichier:', this.audio.src);
        console.log('  - État prêt:', this.audio.readyState);
        console.log('  - Volume:', this.audio.volume);
        console.log('  - Muet:', this.audio.muted);
        
        this.audio.play().then(() => {
            this.isPlaying = true;
            this.updatePlayButtonUI();
            document.getElementById('musicButton').classList.add('playing');
            console.log('✅ Lecture démarrée avec succès!');
            console.log('▶️ En cours:', this.playlist[this.currentTrack].title);
            console.log('🔊 Volume actuel:', Math.round(this.audio.volume * 100) + '%');
        }).catch(error => {
            console.error('❌ Erreur de lecture:', error);
            console.error('  - Nom erreur:', error.name);
            console.error('  - Message:', error.message);
            
            // Essayer de débloquer l'audio avec une interaction utilisateur
            if (error.name === 'NotAllowedError') {
                alert('Cliquez sur OK puis réessayez de lancer la lecture');
                // Créer un contexte audio pour débloquer
                this.unlockAudio();
            } else if (error.name === 'NotSupportedError') {
                alert('Format audio non supporté. Vérifiez que le fichier MP3 est valide.');
            }
        });
    }
    
    unlockAudio() {
        // Créer un contexte audio pour débloquer la lecture
        const AudioContext = window.AudioContext || window.webkitAudioContext;
        if (AudioContext) {
            const audioContext = new AudioContext();
            audioContext.resume().then(() => {
                console.log('🔊 Contexte audio débloqué');
            });
        }
    }
    
    pause() {
        this.audio.pause();
        this.isPlaying = false;
        this.updatePlayButtonUI();
        document.getElementById('musicButton').classList.remove('playing');
    }
    
    previousTrack() {
        this.currentTrack = (this.currentTrack - 1 + this.playlist.length) % this.playlist.length;
        this.loadTrack(this.currentTrack);
        if (this.isPlaying) {
            this.play();
        }
    }
    
    nextTrack() {
        this.currentTrack = (this.currentTrack + 1) % this.playlist.length;
        this.loadTrack(this.currentTrack);
        if (this.isPlaying) {
            this.play();
        }
    }
    
    loadTrack(index) {
        const track = this.playlist[index];
        if (!track) return;
        
        this.currentTrack = index;
        this.audio.src = track.src;
        
        // Log pour debug
        console.log('🎵 Chargement de:', track.title, 'depuis:', track.src);
        
        // Update UI
        document.getElementById('trackTitle').textContent = track.title;
        document.getElementById('trackArtist').textContent = track.artist;
        
        // Update playlist UI
        this.updatePlaylistUI();
        
        // Save preference
        localStorage.setItem('scribe_music_track', index.toString());
    }
    
    seekTo(e) {
        const progressBar = e.currentTarget;
        const clickX = e.offsetX;
        const width = progressBar.offsetWidth;
        const percentage = clickX / width;
        
        this.audio.currentTime = percentage * this.audio.duration;
    }
    
    setVolume(value) {
        this.audio.volume = value;
        this.updateVolumeUI();
        localStorage.setItem('scribe_music_volume', value.toString());
        console.log('🔊 Volume réglé à:', Math.round(value * 100) + '%');
    }
    
    toggleMute() {
        this.audio.muted = !this.audio.muted;
        this.updateVolumeUI();
    }
    
    updateProgress() {
        if (!this.audio.duration) return;
        
        const percentage = (this.audio.currentTime / this.audio.duration) * 100;
        document.getElementById('progressFill').style.width = percentage + '%';
        
        document.getElementById('timeCurrent').textContent = this.formatTime(this.audio.currentTime);
    }
    
    updateDuration() {
        document.getElementById('timeTotal').textContent = this.formatTime(this.audio.duration);
    }
    
    updatePlayButtonUI() {
        const playBtn = document.getElementById('playBtn');
        const playIcon = playBtn.querySelector('.play-icon');
        const pauseIcon = playBtn.querySelector('.pause-icon');
        
        if (this.isPlaying) {
            playIcon.style.display = 'none';
            pauseIcon.style.display = 'block';
        } else {
            playIcon.style.display = 'block';
            pauseIcon.style.display = 'none';
        }
    }
    
    updateVolumeUI() {
        const volumeBtn = document.getElementById('volumeBtn');
        const volumeIcon = volumeBtn.querySelector('.volume-icon');
        const muteIcon = volumeBtn.querySelector('.mute-icon');
        const slider = document.getElementById('volumeSlider');
        
        if (this.audio.muted || this.audio.volume === 0) {
            volumeIcon.style.display = 'none';
            muteIcon.style.display = 'block';
        } else {
            volumeIcon.style.display = 'block';
            muteIcon.style.display = 'none';
        }
        
        slider.value = this.audio.muted ? 0 : this.audio.volume * 100;
    }
    
    updatePlaylistUI() {
        const playlistEl = document.getElementById('playlist');
        playlistEl.innerHTML = '';
        
        this.playlist.forEach((track, index) => {
            const item = document.createElement('div');
            item.className = 'playlist-item';
            if (index === this.currentTrack) {
                item.classList.add('active');
            }
            item.dataset.index = index;
            
            item.innerHTML = `
                <div>
                    <div class="playlist-item-title">${track.title}</div>
                    <div class="playlist-item-artist" style="font-size: 12px; color: #9ca3af;">${track.artist}</div>
                </div>
                <div class="playlist-item-duration">${track.duration}</div>
            `;
            
            playlistEl.appendChild(item);
        });
    }
    
    formatTime(seconds) {
        if (isNaN(seconds)) return '0:00';
        
        const minutes = Math.floor(seconds / 60);
        const secs = Math.floor(seconds % 60);
        
        return `${minutes}:${secs.toString().padStart(2, '0')}`;
    }
    
    destroy() {
        // Nettoyer
        this.pause();
        document.getElementById('musicButton')?.remove();
        document.getElementById('musicPlayer')?.remove();
    }
}

// Export de la classe seulement, pas d'initialisation automatique
// L'initialisation se fera dans app.js
